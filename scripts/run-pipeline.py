#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Pipeline Orchestrator with Extractor Selection Router.

Step 0: Detect format + scan status, select extractor (router)
Step 1: Run convert-paper.py (or alternative extractor script)
Step 1b: Cross-validate extraction with pdfplumber (PDF only)
Step 2: qc-structural.py (GATE: must PASS before proceeding)
Step 3: prepare-image-analysis.py (if images exist)
Step 6a: extract-numbers.py (if PDF format)
Step 6c: Image Index Generation (R19)
Step 3b: Re-run prepare-image-analysis.py (if Step 3 was skipped
         because image-manifest.json did not yet exist at Step 3 time)
Then reports which Claude subagent steps to run.

Usage:
    python3 run-pipeline.py <input_file> [-o output.md] [-i images/] [-s short-name]
    python3 run-pipeline.py paper.pdf --force-extractor tesseract
"""

import argparse
import fcntl
import hashlib
import json
import os
import re
import shlex
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Literal, Optional

# Optional: pdfplumber for cross-validation
try:
    import pdfplumber
    _HAS_PDFPLUMBER = True
except ImportError:
    _HAS_PDFPLUMBER = False

# Optional: fitz (PyMuPDF) for scan detection
try:
    import fitz
    _HAS_FITZ = True
except ImportError:
    _HAS_FITZ = False

# Optional: pymupdf4llm for cross-validation page chunks
try:
    import pymupdf4llm
    _HAS_PYMUPDF4LLM = True
except ImportError:
    _HAS_PYMUPDF4LLM = False

SCRIPTS_DIR = Path(__file__).parent
REGISTRY_PATH = Path.home() / ".claude/pipeline/conversion_registry.json"
SCAN_THRESHOLD = 50  # avg chars/page below this = treat as scanned
PIPELINE_VERSION = "3.2.0"
# Marker wrapper script path (single definition; used by select_extractor,
# _build_cmd_for_extractor, and the runtime fallback chain).
_MARKER_WRAPPER = str(SCRIPTS_DIR / "convert-paper-marker.py")

# ── v3.1 naming convention ────────────────────────────────────────────────
# R9: The subfolder for organized source files is always exactly this name.
# Underscore prefix ensures it sorts first and is visually distinct.
ORIGINALS_SUBDIR = "_originals"

# Heuristic 8: minimum drawing count to classify a page as having
# substantive vector content (SmartArt, shape diagrams, charts).
# Raised from 7 to 50 (BUG-3 fix): styled table headers and colored
# backgrounds produce 20-370 drawings per page, causing false positives
# at the old threshold. Real figures consistently have 140-1,686 drawings.
# Pages with 7-49 drawings are still caught IF they have a drawing whose
# bounding box covers >= 5% of the page area (large shapes = real figures,
# small scattered shapes = table styling).
VECTOR_DRAWING_THRESHOLD = 50

# ── MinerU cross-validation fallback ───────────────────────────────────
# Threshold for auto-switching from pymupdf4llm to MinerU based on
# cross-validation failure rate.  0.40 = 40% of pages flagged with >5%
# word mismatch against pdfplumber.  DOC-1 hit 88% (53/60 pages).
# Minimum 10 pages to avoid false positives on short documents with
# heavy notation (Greek letters, math symbols, units).
MINERU_FALLBACK_THRESHOLD = 0.40
MINERU_FALLBACK_MIN_PAGES = 10

# Path to the MinerU venv Python binary (CPU-only; NEVER use MPS/GPU).
MINERU_PYTHON = Path.home() / "envs" / "mineru" / "bin" / "python3"
MINERU_VENV = Path.home() / "envs" / "mineru"
VECTOR_DRAWING_THRESHOLD_LOW = 7   # Lower bound for area-based check
VECTOR_DRAWING_MIN_AREA_PCT = 5.0  # Min % of page area for a single drawing
VECTOR_DRAWING_HIGH_MIN_AREA_PCT = 1.0  # Min % for >=50 drawings (filters decorative journal styling)

# ── Pipeline issue reporting ───────────────────────────────────────────────
REPORTS_DIR = os.path.join(os.path.expanduser("~"), ".doc2md", "pipeline-improvement", "reports")

# Known failure patterns for auto-detection
KNOWN_FAILURE_PATTERNS = {
    "2000px_crash": {
        "pattern": r"exceeds the dimension limit|2000px|dimension limit for many-image",
        "severity": "CRITICAL",
        "category": "crash",
        "description": "Image exceeds Anthropic 2000px dimension limit for vision API",
        "proposed_fix": "Add image resize step before Opus vision (PIL thumbnail to 2000px max)"
    },
    "blank_image": {
        "pattern": r"blank.*image|image.*blank|empty.*image|zero.*byte",
        "severity": "MAJOR",
        "category": "quality",
        "description": "Blank or empty image detected during extraction",
        "proposed_fix": "Enhance blank detection in extraction step; add 3-tier detection (file size, pixel std, near-black)"
    },
    "missing_manifest": {
        "pattern": r"manifest.*not found|no.*manifest|FileNotFoundError.*manifest",
        "severity": "CRITICAL",
        "category": "crash",
        "description": "Analysis manifest file not found",
        "proposed_fix": "Add manifest existence check before dependent steps; generate skeleton manifest on missing"
    },
    "extraction_failure": {
        "pattern": r"extraction.*fail|failed.*extract|could not extract",
        "severity": "CRITICAL",
        "category": "crash",
        "description": "Image or text extraction failed",
        "proposed_fix": "Add fallback extractor; log specific failure with page number"
    }
}


# ═══════════════════════════════════════════════════════════════════════════
# TYPES
# ═══════════════════════════════════════════════════════════════════════════

ExtractorType = Literal[
    "pymupdf4llm", "tesseract", "mineru", "zerox", "markitdown", "calibre",
    "docling", "marker"
]


@dataclass
class ExtractorConfig:
    """Result of the extractor selection router."""
    extractor: str  # ExtractorType
    script: str     # absolute path to the script to call
    extra_args: list = field(default_factory=list)
    is_scanned: bool = False
    avg_chars_per_page: float = 0.0
    page_count: int = 0


# ═══════════════════════════════════════════════════════════════════════════
# LOGGING
# ═══════════════════════════════════════════════════════════════════════════

def _warn(msg: str) -> None:
    """Log a warning to stderr and continue."""
    print(f"WARN [router]: {msg}", file=sys.stderr)


def _ensure_max_dimension(img_path, max_dim=2000):
    """Resize image in-place if either dimension exceeds max_dim.

    Uses PIL thumbnail() which preserves aspect ratio and uses LANCZOS
    resampling for quality.  Images already within limits are untouched.

    Returns True if image was resized, False otherwise.
    """
    try:
        from PIL import Image as _PilResize
        with _PilResize.open(img_path) as _img:
            if max(_img.size) <= max_dim:
                return False
            _orig = _img.size
            _img.thumbnail((max_dim, max_dim), _PilResize.LANCZOS)
            _img.save(img_path)
            print(f"    [resize] {Path(img_path).name}: "
                  f"{_orig[0]}x{_orig[1]} -> "
                  f"{_img.size[0]}x{_img.size[1]}")
            return True
    except Exception as _e:
        _warn(f"Could not resize {Path(img_path).name}: {_e}")
        return False


def _fail(msg: str, checkpoint: Optional[dict] = None,
          checkpoint_path: Optional[Path] = None) -> None:
    """Log failure, write checkpoint if available, exit 2."""
    print(f"FAIL [router]: {msg}", file=sys.stderr)
    if checkpoint is not None and checkpoint_path is not None:
        checkpoint["current_state"] = "failed"
        checkpoint["failure_reason"] = msg
        try:
            checkpoint_path.parent.mkdir(parents=True, exist_ok=True)
            checkpoint_path.write_text(json.dumps(checkpoint, indent=2))
        except Exception as e:
            print(f"WARN [router]: could not write checkpoint: {e}",
                  file=sys.stderr)
    sys.exit(2)


# ═══════════════════════════════════════════════════════════════════════════
# PIPELINE ISSUE REPORTING
# ═══════════════════════════════════════════════════════════════════════════

def write_pipeline_report(severity, category, description, root_cause="Unknown",
                          impact="Not assessed", proposed_fix="None",
                          affected_files=None, session="auto-detected",
                          auto_detected=False):
    """Write a structured issue report to the pipeline improvement reports folder."""
    import glob as _glob

    os.makedirs(REPORTS_DIR, exist_ok=True)

    date_str = datetime.now().strftime("%Y-%m-%d")
    # Create a short slug from description
    slug = description[:50].lower()
    slug = re.sub(r'[^a-z0-9]+', '-', slug).strip('-')

    # Find unique filename
    base_name = f"{date_str}-{slug}"
    file_path = os.path.join(REPORTS_DIR, f"{base_name}.md")
    counter = 1
    while os.path.exists(file_path):
        file_path = os.path.join(REPORTS_DIR, f"{base_name}-{counter}.md")
        counter += 1

    report_content = f"""# {description}

## Issues

### Issue 1: {description}

- **Session**: {session}
- **Severity**: {severity}
- **Category**: {category}
- **Description**: {description}
- **Root cause**: {root_cause}
- **Impact**: {impact}
- **Proposed fix**: {proposed_fix}
- **Affected files**: {', '.join(affected_files) if affected_files else 'Not specified'}
"""

    with open(file_path, 'w') as f:
        f.write(report_content)

    print(f"[REPORT] Pipeline issue logged to: {file_path}")
    return file_path


def check_for_known_failures(error_message, context=""):
    """Check an error message against known failure patterns and auto-report."""
    for failure_id, failure_info in KNOWN_FAILURE_PATTERNS.items():
        if re.search(failure_info["pattern"], error_message, re.IGNORECASE):
            write_pipeline_report(
                severity=failure_info["severity"],
                category=failure_info["category"],
                description=failure_info["description"],
                root_cause=f"Auto-detected pattern: {failure_id}. Error: {error_message[:200]}",
                impact="Pipeline execution interrupted",
                proposed_fix=failure_info["proposed_fix"],
                affected_files=["scripts/run-pipeline.py"],
                auto_detected=True
            )
            return True
    return False


def run_health_check():
    """Read all pipeline reports and display unresolved issues."""
    import glob as _glob

    report_files = sorted(_glob.glob(os.path.join(REPORTS_DIR, "*.md")))
    report_files = [f for f in report_files if not f.endswith("README.md")
                    and not f.endswith("CROSS-PROJECT-PROMPTS.md")]

    if not report_files:
        print("[HEALTH] No pipeline issue reports found.")
        return

    print(f"\n{'='*60}")
    print(f"PIPELINE HEALTH CHECK — {len(report_files)} report(s) found")
    print(f"{'='*60}\n")

    severity_counts = {"CRITICAL": 0, "MAJOR": 0, "MINOR": 0}

    for report_path in report_files:
        filename = os.path.basename(report_path)
        with open(report_path, 'r') as f:
            content = f.read()

        # Check if resolved
        is_resolved = "RESOLVED:" in content

        # Count severities
        for sev in severity_counts:
            severity_counts[sev] += content.count(f"**Severity**: {sev}")

        status = "RESOLVED" if is_resolved else "OPEN"
        print(f"  [{status}] {filename}")

    print(f"\n{'─'*40}")
    print(f"  CRITICAL: {severity_counts['CRITICAL']}")
    print(f"  MAJOR:    {severity_counts['MAJOR']}")
    print(f"  MINOR:    {severity_counts['MINOR']}")
    print(f"{'─'*40}")

    def _file_is_unresolved(path):
        with open(path, 'r') as fh:
            return "RESOLVED:" not in fh.read()

    open_count = sum(1 for f in report_files if _file_is_unresolved(f))
    if open_count > 0:
        print(f"\n  ⚠ {open_count} OPEN issue(s) need attention")
    else:
        print(f"\n  ✓ All issues resolved")
    print()


def interactive_report_issue():
    """Interactively create a pipeline issue report."""
    print("\n--- Pipeline Issue Report ---\n")
    print(f"Report will be saved to: {REPORTS_DIR}\n")

    session = input("Session (project + session ID): ").strip() or "manual"
    severity = input("Severity (CRITICAL/MAJOR/MINOR): ").strip().upper() or "MINOR"
    category = input("Category (crash/performance/quality/workflow/missing-feature/documentation): ").strip() or "quality"
    description = input("Description: ").strip() or "Unspecified issue"
    root_cause = input("Root cause (or press Enter to skip): ").strip() or "Unknown"
    impact = input("Impact (or press Enter to skip): ").strip() or "Not assessed"
    proposed_fix = input("Proposed fix (or press Enter to skip): ").strip() or "None"
    affected = input("Affected files (comma-separated, or press Enter to skip): ").strip()
    affected_files = [f.strip() for f in affected.split(",")] if affected else None

    path = write_pipeline_report(
        severity=severity, category=category, description=description,
        root_cause=root_cause, impact=impact, proposed_fix=proposed_fix,
        affected_files=affected_files, session=session
    )
    print(f"\nReport saved to: {path}\n")


# ═══════════════════════════════════════════════════════════════════════════
# VECTOR CONTENT HEURISTIC
# ═══════════════════════════════════════════════════════════════════════════

def _has_significant_vector_content(page_data: dict) -> bool:
    """Check whether a page has significant vector content worth rendering.

    Combined heuristic (BUG-3 fix + BUG-2 fix):
      - drawing_count >= 50 (VECTOR_DRAWING_THRESHOLD) AND max_drawing_area_pct
        >= 1.0%: significant.  Real figures have large drawings (5-39% of page).
        NEJM-style journal decorations (column rules, header lines) produce
        50-370 tiny drawings with max_area ~0.0% — the 1.0% guard filters these.
      - drawing_count >= 7 (VECTOR_DRAWING_THRESHOLD_LOW) AND the page has
        at least one drawing whose bounding box covers >= 5% of the page
        area: significant.  This catches simple vector figures with few
        but large shapes, while filtering out table styling (many small
        scattered drawings like colored cell backgrounds).
    """
    dc = page_data.get("drawing_count", 0)
    if dc >= VECTOR_DRAWING_THRESHOLD:
        area_pct = page_data.get("max_drawing_area_pct", 0.0)
        if area_pct >= VECTOR_DRAWING_HIGH_MIN_AREA_PCT:
            return True
        # Many tiny decorative drawings (journal styling) — log for diagnostics
        page_num = page_data.get("page", "?")
        print(f"  Vector: page {page_num} has {dc} drawings but "
              f"max_area={area_pct:.1f}% < "
              f"{VECTOR_DRAWING_HIGH_MIN_AREA_PCT}% "
              f"— skipping (decorative styling)")
        return False
    if dc >= VECTOR_DRAWING_THRESHOLD_LOW:
        area_pct = page_data.get("max_drawing_area_pct", 0.0)
        if area_pct >= VECTOR_DRAWING_MIN_AREA_PCT:
            return True
    return False


# ═══════════════════════════════════════════════════════════════════════════
# SCAN DETECTION
# ═══════════════════════════════════════════════════════════════════════════

def _measure_text_density(pdf_path: Path) -> tuple:
    """Return (avg_chars_per_page, page_count).

    Samples up to 10 evenly-spaced pages for speed.
    Returns (100.0, 0) if fitz is unavailable or file cannot be opened
    (assumes digital to allow pymupdf4llm to try).
    """
    if not _HAS_FITZ:
        _warn("fitz (PyMuPDF) not installed. Cannot detect scan status. "
              "Assuming digital PDF.")
        return (100.0, 0)

    try:
        doc = fitz.open(str(pdf_path))
    except Exception as e:
        _warn(f"Could not open PDF for scan detection: {e}. "
              "Assuming digital PDF.")
        return (100.0, 0)

    page_count = len(doc)
    if page_count == 0:
        doc.close()
        return (0.0, 0)

    # Sample up to 10 evenly-spaced pages
    step = max(1, page_count // 10)
    sample_indices = list(range(0, page_count, step))
    total_chars = 0
    for i in sample_indices:
        try:
            total_chars += len(doc[i].get_text("text"))
        except Exception:
            pass  # skip unreadable pages in the sample

    doc.close()
    avg = total_chars / len(sample_indices) if sample_indices else 0.0
    return (avg, page_count)


# ═══════════════════════════════════════════════════════════════════════════
# AVAILABILITY CHECKS
# ═══════════════════════════════════════════════════════════════════════════

def _tesseract_available() -> bool:
    """Check if tesseract binary is on PATH."""
    return shutil.which("tesseract") is not None


def _mineru_available() -> bool:
    """Check if MinerU venv exists at ~/envs/mineru/ with a valid Python binary."""
    python_bin = Path.home() / "envs" / "mineru" / "bin" / "python3"
    return python_bin.exists()


def _get_mineru_version() -> str:
    """Detect the installed MinerU (magic-pdf) version at runtime.

    Tries importlib.metadata first (fast), then falls back to
    pip show in the MinerU venv.  Returns 'unknown' if detection fails.
    """
    # Try importlib.metadata (works if magic-pdf is importable)
    try:
        import importlib.metadata
        return importlib.metadata.version("magic-pdf")
    except Exception:
        pass
    # Fallback: ask pip inside the MinerU venv
    pip_bin = MINERU_VENV / "bin" / "pip"
    if pip_bin.exists():
        try:
            result = subprocess.run(
                [str(pip_bin), "show", "magic-pdf"],
                capture_output=True, text=True, timeout=15,
            )
            for line in result.stdout.splitlines():
                if line.startswith("Version:"):
                    return line.split(":", 1)[1].strip()
        except Exception:
            pass
    return "unknown"


def _zerox_available() -> bool:
    """Check if zerox convert script exists."""
    return (SCRIPTS_DIR / "convert-zerox.py").exists()


def _docling_available() -> bool:
    """Check if docling is importable in the current Python environment."""
    try:
        from docling.document_converter import DocumentConverter  # noqa: F401
        return True
    except ImportError:
        return False


def _pymupdf4llm_available() -> bool:
    """Check if pymupdf4llm extractor is available.

    Wraps the module-level _HAS_PYMUPDF4LLM flag for consistency with
    _docling_available(), _mineru_available(), _tesseract_available().
    """
    return _HAS_PYMUPDF4LLM


def _marker_available() -> bool:
    """Check if marker-pdf CLI (marker_single) is available."""
    return shutil.which("marker_single") is not None


# ═══════════════════════════════════════════════════════════════════════════
# EXTRACTOR SELECTION ROUTER
# ═══════════════════════════════════════════════════════════════════════════

def select_extractor(pdf_path: Path,
                     force_extractor: Optional[str] = None) -> ExtractorConfig:
    """Detect scan status and return the correct extractor config.

    Called ONLY for PDF files. Non-PDF formats bypass this function.

    Chains:
      Digital (>= 50 chars/page):
        marker -> docling -> pymupdf4llm -> mineru -> tesseract
        NOTE: Digital table fallbacks (Camelot -> pdfplumber -> MinerU)
        are triggered by qc-structural.py downstream, not by this router.
      Scanned (< 50 chars/page):
        Tesseract -> MinerU -> Zerox -> flag for manual review
        NOTE: This is availability-based selection. Runtime failure
        fallback is handled in main() via _build_cmd_for_extractor()
        and _next_scanned_fallback().

    Args:
        pdf_path: Path to the PDF file.
        force_extractor: Override automatic detection.

    Returns:
        ExtractorConfig with the selected extractor and invocation details.
    """
    convert_paper = str(SCRIPTS_DIR / "convert-paper.py")
    convert_mineru = str(SCRIPTS_DIR / "convert-mineru.py")

    # ── Measure text density ONCE (Issue 7: avoid double PDF scan) ──
    avg, page_count = _measure_text_density(pdf_path)
    is_scanned = avg < SCAN_THRESHOLD

    # ── Issue 2: Guard against 0-page PDF ──
    if page_count == 0 and _HAS_FITZ:
        _warn("PDF has 0 pages. Cannot extract content. "
              "Flagging for manual review.")
        # Return pymupdf4llm as best-effort; will produce empty output
        # which downstream steps will detect and flag
        return ExtractorConfig(
            extractor="pymupdf4llm",
            script=convert_paper,
            extra_args=["--extractor", "pymupdf4llm"],
            is_scanned=False,
            avg_chars_per_page=0.0,
            page_count=0,
        )

    # ── Force override ──
    if force_extractor:
        _warn(f"Forcing extractor '{force_extractor}' "
              "-- skipping availability check and auto-detection")
        if force_extractor == "tesseract":
            return ExtractorConfig(
                extractor="tesseract",
                script=convert_paper,
                extra_args=["--extractor", "tesseract"],
                is_scanned=is_scanned,
                avg_chars_per_page=avg,
                page_count=page_count,
            )
        elif force_extractor == "mineru":
            return ExtractorConfig(
                extractor="mineru",
                script=convert_mineru,
                extra_args=[],
                is_scanned=is_scanned,
                avg_chars_per_page=avg,
                page_count=page_count,
            )
        elif force_extractor == "pymupdf4llm":
            return ExtractorConfig(
                extractor="pymupdf4llm",
                script=convert_paper,
                extra_args=["--extractor", "pymupdf4llm"],
                is_scanned=is_scanned,
                avg_chars_per_page=avg,
                page_count=page_count,
            )
        elif force_extractor == "docling":
            return ExtractorConfig(
                extractor="docling",
                script=convert_paper,
                extra_args=["--extractor", "docling"],
                is_scanned=is_scanned,
                avg_chars_per_page=avg,
                page_count=page_count,
            )
        elif force_extractor == "marker":
            return ExtractorConfig(
                extractor="marker",
                script=_MARKER_WRAPPER,
                extra_args=[],
                is_scanned=is_scanned,
                avg_chars_per_page=avg,
                page_count=page_count,
            )
        else:
            _warn(f"Unknown forced extractor '{force_extractor}'. "
                  "Falling through to auto-detection.")

    # ── Auto-detection (uses pre-computed avg and is_scanned) ──
    print(f"  Scan detection: avg {avg:.1f} chars/page "
          f"({page_count} pages) -> "
          f"{'SCANNED' if is_scanned else 'DIGITAL'}")

    if not is_scanned:
        # ── Digital chain ──
        # marker is the new default for digital PDFs (S42). Falls back
        # to docling then pymupdf4llm via _DIGITAL_CHAIN.
        if _marker_available():
            return ExtractorConfig(
                extractor="marker",
                script=_MARKER_WRAPPER,
                extra_args=[],
                is_scanned=False,
                avg_chars_per_page=avg,
                page_count=page_count,
            )
        # marker unavailable -> try docling
        if _docling_available():
            _warn("Marker not available. Falling back to docling.")
            return ExtractorConfig(
                extractor="docling",
                script=convert_paper,
                extra_args=["--extractor", "docling"],
                is_scanned=False,
                avg_chars_per_page=avg,
                page_count=page_count,
            )
        # Docling also unavailable -> fall back to pymupdf4llm
        _warn("Marker and docling not available. "
              "Falling back to pymupdf4llm.")
        return ExtractorConfig(
            extractor="pymupdf4llm",
            script=convert_paper,
            extra_args=["--extractor", "pymupdf4llm"],
            is_scanned=False,
            avg_chars_per_page=avg,
            page_count=page_count,
        )

    # ── Scanned chain ──
    # Try Tesseract first (fast, zero extra dependency if PyMuPDF present)
    if _tesseract_available():
        print("  Scanned -> selecting Tesseract OCR via PyMuPDF")
        return ExtractorConfig(
            extractor="tesseract",
            script=convert_paper,
            extra_args=["--extractor", "tesseract"],
            is_scanned=True,
            avg_chars_per_page=avg,
            page_count=page_count,
        )
    else:
        _warn("Tesseract not installed. Trying MinerU.")

    # Tesseract unavailable -> MinerU
    if _mineru_available():
        print("  Scanned -> selecting MinerU")
        return ExtractorConfig(
            extractor="mineru",
            script=convert_mineru,
            extra_args=[],
            is_scanned=True,
            avg_chars_per_page=avg,
            page_count=page_count,
        )
    else:
        _warn("MinerU venv not found at ~/envs/mineru/. Trying Zerox.")

    # MinerU unavailable -> Zerox (VLM, last resort)
    if _zerox_available():
        print("  Scanned -> selecting Zerox VLM (last resort)")
        return ExtractorConfig(
            extractor="zerox",
            script=str(SCRIPTS_DIR / "convert-zerox.py"),
            extra_args=[],
            is_scanned=True,
            avg_chars_per_page=avg,
            page_count=page_count,
        )
    else:
        _warn("Zerox not available. No OCR extractor found.")

    # Nothing available -> flag for manual review
    _warn("ALL scanned-PDF extractors unavailable. "
          "Flagging for manual review. Attempting pymupdf4llm "
          "as best-effort fallback.")
    return ExtractorConfig(
        extractor="pymupdf4llm",
        script=convert_paper,
        extra_args=["--extractor", "pymupdf4llm"],
        is_scanned=True,
        avg_chars_per_page=avg,
        page_count=page_count,
    )


# ═══════════════════════════════════════════════════════════════════════════
# CROSS-VALIDATION
# ═══════════════════════════════════════════════════════════════════════════

def cross_validate_extraction(pdf_path: Path,
                              md_chunks: list) -> list:
    """Cross-validate pymupdf4llm output against pdfplumber.

    Returns list of flagged page dicts:
        {page, completeness, missing_sample}
    Caller writes flags to checkpoint; qc-structural.py reads them.

    Args:
        pdf_path: Path to the original PDF.
        md_chunks: list of dicts from pymupdf4llm page_chunks=True,
                   each with a "text" key.
    """
    if not _HAS_PDFPLUMBER:
        _warn("pdfplumber not installed. Cross-validation skipped.")
        return []

    flags = []
    try:
        with pdfplumber.open(str(pdf_path)) as pdf:
            for i, page in enumerate(pdf.pages):
                if i >= len(md_chunks):
                    break
                secondary_text = page.extract_text() or ""
                primary_text = md_chunks[i].get("text", "")

                primary_words = set(primary_text.split())
                secondary_words = set(secondary_text.split())

                if not secondary_words:
                    continue  # pdfplumber got nothing, skip

                missing = secondary_words - primary_words
                ratio = len(missing) / len(secondary_words)

                if ratio > 0.05:
                    flags.append({
                        "page": i,
                        "completeness": round(1 - ratio, 4),
                        "missing_sample": list(missing)[:10],
                    })
    except Exception as e:
        _warn(f"Cross-validation failed: {e}")

    return flags


def _get_page_chunks(pdf_path: Path) -> list:
    """Get per-page text chunks using pymupdf4llm for cross-validation.

    Returns list of dicts with "text" key, one per page.
    Returns empty list if pymupdf4llm is unavailable.
    """
    if not _HAS_PYMUPDF4LLM:
        _warn("pymupdf4llm not available for cross-validation chunks.")
        return []

    try:
        chunks = pymupdf4llm.to_markdown(
            str(pdf_path),
            page_chunks=True,
            show_progress=False,
            write_images=False,
        )
        return chunks
    except Exception as e:
        _warn(f"Could not get page chunks for cross-validation: {e}")
        return []


def _is_slide_based_pdf(pdf_path: Path,
                        md_chunks: list = None) -> bool:
    """Detect if a PDF is a presentation-style (slide-based) document.

    Uses heuristics: landscape orientation, consistent page sizes,
    low text density, high image-to-text ratio.
    Returns True if the document appears to be a converted presentation.

    Args:
        pdf_path: Path to the original PDF file.
        md_chunks: Optional list of per-page text dicts from pymupdf4llm.
    """
    if not _HAS_FITZ:
        return False

    try:
        doc = fitz.open(str(pdf_path))
    except Exception:
        return False

    if len(doc) < 3:
        doc.close()
        return False

    # ── Heuristic 1: Landscape orientation (majority of pages) ──
    landscape_count = 0
    for page in doc:
        rect = page.rect
        if rect.width > rect.height:
            landscape_count += 1
    landscape_ratio = landscape_count / len(doc)

    # ── Heuristic 2: Consistent page sizes (presentation = uniform) ──
    page_sizes = set()
    for page in doc:
        rect = page.rect
        # Round to nearest 10 to allow minor float differences
        page_sizes.add((round(rect.width, -1), round(rect.height, -1)))
    uniform_pages = len(page_sizes) <= 2  # 1 or 2 unique sizes

    # ── Heuristic 3: Low text density ──
    low_text_pages = 0
    if md_chunks:
        for chunk in md_chunks:
            text = chunk.get("text", "")
            word_count = len(text.split())
            if word_count < 80:  # Slides typically have < 80 words
                low_text_pages += 1
        low_text_ratio = low_text_pages / len(md_chunks) if md_chunks else 0
    else:
        # Fallback: use fitz text extraction
        for page in doc:
            text = page.get_text()
            if len(text.split()) < 80:
                low_text_pages += 1
        low_text_ratio = low_text_pages / len(doc)

    # ── Heuristic 4: Aspect ratio check (16:9 or 4:3) ──
    common_ratios = 0
    for page in doc:
        rect = page.rect
        if rect.height > 0:
            ratio = rect.width / rect.height
            # 16:9 ≈ 1.78, 4:3 ≈ 1.33
            if (1.2 < ratio < 1.5) or (1.6 < ratio < 2.0):
                common_ratios += 1
    ratio_match = common_ratios / len(doc) if len(doc) > 0 else 0

    doc.close()

    # ── Decision: need at least 2 of 4 heuristics to trigger ──
    # H2 (uniform_pages) alone cannot distinguish slides from academic
    # papers — nearly every professional document has consistent page
    # sizes.  Require co-occurrence with landscape (H1) or aspect-ratio
    # (H4) to count H2 as a signal.
    if uniform_pages and landscape_ratio < 0.80 and ratio_match < 0.80:
        uniform_pages = False

    signals = 0
    if landscape_ratio >= 0.8:
        signals += 1
    if uniform_pages:
        signals += 1
    if low_text_ratio >= 0.6:
        signals += 1
    if ratio_match >= 0.8:
        signals += 1

    return signals >= 2


# ═══════════════════════════════════════════════════════════════════════════
# DOMAIN DETECTION (F14)
# ═══════════════════════════════════════════════════════════════════════════

# Keyword lists for domain detection.
# SYNC: mirrors DOMAIN_KEYWORDS / DOMAIN_OVERRIDE_KEYWORDS in convert-paper.py
# with additional health_economics keywords from F14 spec.
_DOMAIN_OVERRIDE_KEYWORDS = {
    "hta_regulatory": [
        "NoMA", "metodevurdering", "ICER threshold", "cost per QALY",
        "NICE appraisal", "health technology assessment",
        "reimbursement decision", "HTA body", "DMP",
    ],
    "health_economics": [
        "Markov model", "cost-effectiveness analysis",
        "willingness to pay", "ICER",
        "incremental cost-effectiveness",
        "cost-utility analysis", "survival analysis",
        "hazard ratio",
    ],
}

_DOMAIN_KEYWORDS = {
    "health_economics": [
        "cost", "QALY", "ICER", "Markov", "willingness-to-pay",
        "cost-effectiveness", "incremental", "budget impact", "threshold",
        "cost-utility", "quality-adjusted life year",
        "cost-benefit", "health economics", "economic evaluation",
        "survival analysis", "hazard ratio", "Kaplan-Meier",
        "probabilistic sensitivity", "tornado diagram",
        "acceptability curve", "net monetary benefit",
    ],
    "clinical_trial": [
        "randomized", "placebo", "endpoint", "ITT", "CONSORT",
        "adverse event", "RCT", "trial", "phase", "randomization",
    ],
    "systematic_review": [
        "PRISMA", "meta-analysis", "forest plot", "I-squared",
        "pooled", "heterogeneity", "systematic", "review",
    ],
    "hta_regulatory": [
        "NICE", "DMP", "NoMA", "reimbursement", "submission",
        "appraisal", "HTA", "regulatory", "guideline",
    ],
    "epidemiology": [
        "incidence", "prevalence", "registry", "cohort", "DALY",
        "mortality", "population", "burden", "disease",
    ],
    "methodology": [
        "model validation", "simulation", "calibration",
        "structural uncertainty", "sensitivity", "probabilistic",
    ],
    "pharmaceutical": [
        "pharmacokinetics", "pharmacodynamics", "bioequivalence",
        "absorption", "clearance", "half-life", "AUC",
        "Cmax", "bioavailability", "formulation",
    ],
}


def _detect_document_domain(md_text: str) -> tuple:
    """Auto-detect document domain from keyword frequencies.

    Returns (domain: str, keyword_count: int, matched_keywords: list).
    domain is one of: health_economics, clinical_trial, systematic_review,
    hta_regulatory, epidemiology, methodology, pharmaceutical, general.

    F14: Enhanced with additional health_economics keywords (ICER, QALY,
    cost-effectiveness, cost-utility, survival analysis, hazard ratio, etc.)
    and keyword count logging for verification.

    S21/RC1: health_economics overrides take priority over hta_regulatory
    when BOTH have override matches.  hta_regulatory only wins in Phase 1
    when health_economics has zero override matches (i.e. the document is
    about regulatory process/HTA methodology without economic evaluation).
    """
    text_lower = md_text.lower()

    # ── Phase 1: Override keywords (high-specificity terms) ──
    # Collect ALL override matches first, then resolve priority.
    override_hits = {}  # domain -> list of matched keywords
    for domain, override_kws in _DOMAIN_OVERRIDE_KEYWORDS.items():
        matched = [kw for kw in override_kws if kw.lower() in text_lower]
        if matched:
            override_hits[domain] = matched

    if override_hits:
        # S21/RC1: health_economics always wins over hta_regulatory when
        # both have override hits (documents about cost-effectiveness that
        # also mention HTA bodies are health_economics, not regulatory).
        if "health_economics" in override_hits:
            m = override_hits["health_economics"]
            return ("health_economics", len(m), m)
        # Only one domain matched, or hta_regulatory without health_econ
        best_domain = max(override_hits, key=lambda d: len(override_hits[d]))
        m = override_hits[best_domain]
        return (best_domain, len(m), m)

    # ── Phase 2: Frequency scoring ──
    domain_scores = {}
    domain_matched = {}
    for domain, keywords in _DOMAIN_KEYWORDS.items():
        matched = [kw for kw in keywords if kw.lower() in text_lower]
        domain_scores[domain] = len(matched)
        domain_matched[domain] = matched

    # Return domain with highest score, or "general" if all zero
    max_domain = max(domain_scores.items(), key=lambda x: x[1])
    if max_domain[1] == 0:
        return ("general", 0, [])

    best = max_domain[0]
    return (best, domain_scores[best], domain_matched[best])


# ═══════════════════════════════════════════════════════════════════════════
# REGISTRY
# ═══════════════════════════════════════════════════════════════════════════

def _compute_sha256(file_path: Path) -> str:
    """Compute SHA-256 hash of a file."""
    h = hashlib.sha256()
    with open(file_path, "rb") as f:
        while True:
            chunk = f.read(8192)
            if not chunk:
                break
            h.update(chunk)
    return f"sha256:{h.hexdigest()}"


def update_registry(pdf_path: Path, output_md: Path,
                    source_hash: str, extractor: str,
                    image_index_meta: Optional[dict] = None) -> None:
    """Append entry to conversion_registry.json.

    Deduplicates by source_hash (replaces existing entry for same hash).
    Registry write failure is WARN only - does not abort pipeline.

    R21: When image_index_meta is provided, adds 7 image-related fields:
        image_index_path, image_index_generated_at, total_pages,
        pages_with_images, total_images_detected, substantive_images,
        has_testable_images.
    """
    REGISTRY_PATH.parent.mkdir(parents=True, exist_ok=True)

    # Lock the entire read-modify-write cycle to prevent concurrent
    # lost updates (MINOR-2 fix).
    # NOTE: The .json.lock file is intentionally persistent (never deleted).
    # fcntl.flock() works on existing files, and keeping the file avoids a
    # race condition where two processes try to create it simultaneously.
    _lock_path = REGISTRY_PATH.with_suffix(".json.lock")
    with open(_lock_path, "w") as _lf:
        fcntl.flock(_lf, fcntl.LOCK_EX)
        try:
            registry = {"conversions": []}
            if REGISTRY_PATH.exists():
                try:
                    with open(REGISTRY_PATH) as f:
                        registry = json.load(f)
                    if not isinstance(registry, dict) or "conversions" not in registry:
                        raise ValueError("malformed registry")
                except (json.JSONDecodeError, ValueError) as e:
                    backup = REGISTRY_PATH.with_suffix(".json.bak")
                    _warn(f"Registry corrupt ({e}). Backing up to {backup} "
                          "and starting fresh.")
                    try:
                        shutil.copy2(REGISTRY_PATH, backup)
                    except Exception:
                        pass
                    registry = {"conversions": []}

            # Remove existing entry with same hash (dedup)
            existing = [c for c in registry.get("conversions", [])
                        if c.get("source_hash") != source_hash]
            new_entry = {
                "source_hash": source_hash,
                "source_path": str(pdf_path.resolve()),
                "output_path": str(output_md.resolve()),
                "extractor_used": extractor,
                "converted_at": datetime.now().isoformat(),
                "pipeline_version": PIPELINE_VERSION,
            }

            # R21: Add image index metadata if available
            if image_index_meta is not None:
                new_entry["image_index_path"] = image_index_meta.get(
                    "image_index_path", "")
                new_entry["image_index_generated_at"] = image_index_meta.get(
                    "image_index_generated_at", "")
                new_entry["total_pages"] = image_index_meta.get(
                    "total_pages", 0)
                new_entry["pages_with_images"] = image_index_meta.get(
                    "pages_with_images", 0)
                new_entry["total_images_detected"] = image_index_meta.get(
                    "total_images_detected", 0)
                new_entry["substantive_images"] = image_index_meta.get(
                    "substantive_images", 0)
                new_entry["has_testable_images"] = image_index_meta.get(
                    "has_testable_images", False)

            existing.append(new_entry)
            registry["conversions"] = existing

            # Atomic write: temp file then replace
            tmp_path = REGISTRY_PATH.with_suffix(".json.tmp")
            with open(tmp_path, "w") as f:
                json.dump(registry, f, indent=2)
            tmp_path.replace(REGISTRY_PATH)
        finally:
            fcntl.flock(_lf, fcntl.LOCK_UN)

    print(f"\nRegistry updated: {REGISTRY_PATH}")
    print(f"  source_hash: {source_hash}")
    print(f"  extractor:   {extractor}")
    if image_index_meta:
        print(f"  image_index_path: "
              f"{new_entry.get('image_index_path', 'N/A')}")
        print(f"  has_testable_images: "
              f"{new_entry.get('has_testable_images', False)}")


# ═══════════════════════════════════════════════════════════════════════════
# COMMAND RUNNER
# ═══════════════════════════════════════════════════════════════════════════

def run_command(cmd: list, description: str,
                allow_failure: bool = False,
                timeout: Optional[int] = None) -> int:
    """Run a command and report status. Returns exit code.

    Args:
        cmd: Command list for subprocess.run.
        description: Human-readable description for logging.
        allow_failure: If True, do not log ERROR on nonzero exit.
        timeout: Optional timeout in seconds. Returns exit code 124
                 (consistent with GNU timeout) if exceeded.
    """
    print(f"\n{'=' * 60}")
    print(f"Running: {description}")
    print(f"Command: {shlex.join(str(c) for c in cmd)}")
    if timeout is not None:
        print(f"Timeout: {timeout}s")
    print('=' * 60)

    try:
        result = subprocess.run(cmd, capture_output=False,
                                timeout=timeout)
    except subprocess.TimeoutExpired:
        print(f"\nERROR: {description} timed out after {timeout}s")
        return 124  # GNU timeout convention

    if result.returncode != 0 and not allow_failure:
        print(f"\nERROR: {description} failed with exit code "
              f"{result.returncode}")
        return result.returncode

    return result.returncode


# ═══════════════════════════════════════════════════════════════════════════
# EXTRACTOR QUALITY GATE (shared helper)
# ═══════════════════════════════════════════════════════════════════════════

def _extractor_quality_gate(extractor_name: str, output_md: Path,
                            input_pdf: Path) -> bool:
    """Lightweight word-count quality gate: compare extractor output vs fitz.

    Used by docling and marker quality gates after extraction to flag
    potentially incomplete conversions.  Logs a WARNING if the extractor
    produced <30% of the fitz word count, INFO if <60%, and INFO-OK
    otherwise.

    Returns True if extraction is critically empty (<10% of expected
    text), signaling the caller to treat this as a failed extraction
    and trigger the fallback chain.  Returns False otherwise.

    Args:
        extractor_name: Human-readable extractor name (e.g. "Docling", "Marker").
        output_md: Path to the extractor's markdown output.
        input_pdf: Path to the source PDF (read by fitz for reference text).
    """
    try:
        import fitz as _fitz_qg
        _doc_qg = _fitz_qg.open(str(input_pdf))
        _fitz_text = ""
        for _page in _doc_qg:
            _fitz_text += _page.get_text("text")
        _doc_qg.close()

        _fitz_words = len(_fitz_text.split())
        _ext_text = output_md.read_text(encoding="utf-8")
        _ext_words = len(_ext_text.split())

        if _fitz_words > 0:
            _ratio = _ext_words / _fitz_words
            # Determine next-in-chain fallback for the warning message
            _fallback_hint = {
                "Docling": "pymupdf4llm",
                "Marker": "docling",
            }.get(extractor_name, "next extractor")
            if _ratio < 0.1:
                _warn(
                    f"{extractor_name} output is critically empty: "
                    f"{_ratio:.0%} of expected text ({_ext_words} vs "
                    f"{_fitz_words} fitz words). Forcing fallback "
                    f"to {_fallback_hint}.")
                return True  # Critical: trigger fallback
            elif _ratio < 0.3:
                _warn(
                    f"{extractor_name} output has only {_ratio:.0%} of "
                    f"expected text ({_ext_words} vs "
                    f"{_fitz_words} fitz words). Consider "
                    f"fallback to {_fallback_hint}.")
            elif _ratio < 0.6:
                print(
                    f"  [INFO] {extractor_name} word ratio: {_ratio:.0%} "
                    f"({_ext_words}/{_fitz_words} fitz words)")
            else:
                print(
                    f"  [INFO] {extractor_name} word ratio: {_ratio:.0%} "
                    f"({_ext_words}/{_fitz_words} fitz words)"
                    f" - OK")
    except Exception as _e_qg:
        print(f"  [WARNING] {extractor_name} quality check failed: {_e_qg}")
    return False


# ═══════════════════════════════════════════════════════════════════════════
# SCANNED PDF RUNTIME FALLBACK
# ═══════════════════════════════════════════════════════════════════════════

# Ordered fallback chain for scanned PDFs.
# select_extractor() picks the first AVAILABLE extractor.
# If that extractor FAILS at runtime, main() tries the next in chain.
_SCANNED_CHAIN = ["tesseract", "mineru", "zerox"]

# Ordered fallback chain for digital PDFs.
# marker is the default (S42); if it crashes, the router tries docling
# then pymupdf4llm then mineru then tesseract before giving up.
_DIGITAL_CHAIN = ["marker", "docling", "pymupdf4llm", "mineru", "tesseract"]


def _build_cmd_for_extractor(extractor: str, input_file: Path,
                              output_md: Path, images_dir: Path,
                              short_name: str,
                              no_images: bool = False) -> list:
    """Build the subprocess command list for a given extractor.

    Centralizes command construction so both the initial run and
    fallback retries use the same logic.
    """
    if extractor in ("pymupdf4llm", "tesseract", "docling"):
        cmd = [
            sys.executable,
            str(SCRIPTS_DIR / "convert-paper.py"),
            str(input_file),
            "-o", str(output_md),
            "-i", str(images_dir),
            "-s", short_name,
            "--extractor", extractor,
        ]
    elif extractor == "marker":
        cmd = [
            sys.executable,
            _MARKER_WRAPPER,
            str(input_file),
            "--output-dir", str(output_md.parent),
        ]
    elif extractor == "mineru":
        cmd = [
            sys.executable,
            str(SCRIPTS_DIR / "convert-mineru.py"),
            str(input_file),
            "--output", str(output_md),
        ]
    elif extractor == "zerox":
        cmd = [
            sys.executable,
            str(SCRIPTS_DIR / "convert-zerox.py"),
            str(input_file),
            "--output", str(output_md),
        ]
    else:
        # Unknown extractor: best-effort via convert-paper.py
        cmd = [
            sys.executable,
            str(SCRIPTS_DIR / "convert-paper.py"),
            str(input_file),
            "-o", str(output_md),
            "-i", str(images_dir),
            "-s", short_name,
        ]
    if no_images and extractor in ("pymupdf4llm", "tesseract", "docling"):
        cmd.append("--no-images")
    return cmd


def _next_scanned_fallback(current_extractor: str) -> Optional[str]:
    """Return the next extractor in the scanned fallback chain.

    Returns None if current_extractor is the last in the chain
    or not in the chain at all.
    """
    try:
        idx = _SCANNED_CHAIN.index(current_extractor)
    except ValueError:
        return None
    # Walk forward to find the next AVAILABLE extractor
    for candidate in _SCANNED_CHAIN[idx + 1:]:
        if candidate == "mineru" and _mineru_available():
            return candidate
        elif candidate == "zerox" and _zerox_available():
            return candidate
        elif candidate == "tesseract" and _tesseract_available():
            return candidate
    return None


def _next_digital_fallback(current_extractor: str) -> Optional[str]:
    """Return the next extractor in the digital fallback chain.

    Used when an extractor crashes on a digital PDF (e.g. ValueError
    in table detection code inside pymupdf/table.py).

    Fallback order: marker -> docling -> pymupdf4llm -> mineru -> tesseract
    Returns None if there is no available next extractor.
    """
    try:
        idx = _DIGITAL_CHAIN.index(current_extractor)
    except ValueError:
        return None
    for candidate in _DIGITAL_CHAIN[idx + 1:]:
        if candidate == "marker" and _marker_available():
            return candidate
        elif candidate == "docling" and _docling_available():
            return candidate
        elif candidate == "pymupdf4llm" and _pymupdf4llm_available():
            return candidate
        elif candidate == "mineru" and _mineru_available():
            return candidate
        elif candidate == "tesseract" and _tesseract_available():
            return candidate
    return None


# ═══════════════════════════════════════════════════════════════════════════
# MINERU CROSS-VALIDATION FALLBACK (Issue 2)
# ═══════════════════════════════════════════════════════════════════════════

def _is_near_black(image_path: Path, threshold: int = 10) -> bool:
    """Check if an image is near-black (overwhelmingly dark pixels).

    Detection strategy (two passes):
        Pass 1 (legacy): mean < threshold AND std < 5.
            Catches uniformly dark/blank images quickly.
        Pass 2 (pixel-percentage): if mean < threshold, count pixels
            with brightness < 15 across all channels.  If > 95% of
            pixels are below that brightness → near-black.
            This catches dark charts with sparse bright features
            (axis labels, data points) where std is high (16-30)
            due to those sparse features but 95%+ pixels are black.

    Returns True if the image is near-black by either pass.
    Returns False on any error (safe default: keep the image).
    """
    try:
        from PIL import Image
        import numpy as np
        img = Image.open(image_path).convert("RGB")
        arr = np.array(img)
        mean_val = float(arr.mean())
        if mean_val >= threshold:
            return False

        # Pass 1: uniform darkness (legacy check).
        std_val = float(arr.std())
        if std_val < 5:
            return True

        # Pass 2: pixel-percentage approach.
        # Convert to grayscale luminance for per-pixel brightness.
        # Using standard luminance weights: 0.299R + 0.587G + 0.114B
        gray = (0.299 * arr[:, :, 0]
                + 0.587 * arr[:, :, 1]
                + 0.114 * arr[:, :, 2])
        total_pixels = gray.size
        dark_pixels = int(np.sum(gray < 15))
        dark_ratio = dark_pixels / total_pixels
        if dark_ratio > 0.95:
            return True

        return False
    except Exception:
        return False


def _rerender_page_pdftoppm(pdf_path: Path, page_num: int,
                            output_path: Path) -> bool:
    """Re-render a single PDF page at 300 DPI using pdftoppm.

    Args:
        pdf_path: Path to the source PDF.
        page_num: 1-based page number.
        output_path: Desired output PNG path.

    Returns:
        True if re-render succeeded, False otherwise.
    """
    if not shutil.which("pdftoppm"):
        _warn("pdftoppm not available for near-black re-render.")
        return False

    with tempfile.TemporaryDirectory(prefix="pdftoppm-") as tmpdir:
        prefix = Path(tmpdir) / "page"
        cmd = [
            "pdftoppm", "-png", "-r", "300",
            "-f", str(page_num), "-l", str(page_num),
            str(pdf_path), str(prefix),
        ]
        try:
            result = subprocess.run(
                cmd, capture_output=True, text=True, timeout=60)
            if result.returncode != 0:
                _warn(f"pdftoppm failed for page {page_num}: "
                      f"{result.stderr[:200]}")
                return False

            # pdftoppm outputs: <prefix>-<pagenum>.png (zero-padded)
            rendered = list(Path(tmpdir).glob("page-*.png"))
            if not rendered:
                _warn(f"pdftoppm produced no output for page {page_num}")
                return False

            shutil.copy2(rendered[0], output_path)
            _ensure_max_dimension(output_path)
            return True
        except subprocess.TimeoutExpired:
            _warn(f"pdftoppm timed out for page {page_num}")
            return False
        except Exception as e:
            _warn(f"pdftoppm error for page {page_num}: {e}")
            return False


def _detect_extraction_gaps(
    pdf_path: Path,
    mineru_images_dir: Path,
    mineru_manifest_images: list,
) -> list:
    """Detect pages where fitz finds images but MinerU extracted none.

    Compares PyMuPDF's per-page image detection (page.get_images()) against
    the MinerU manifest to find "extraction gaps" — pages with embedded
    raster images that MinerU missed.

    Args:
        pdf_path: Path to the original PDF file.
        mineru_images_dir: Path to the pipeline images directory containing
                          MinerU-extracted images.
        mineru_manifest_images: List of manifest image dicts from MinerU
                               normalization (each has a "page" field).

    Returns:
        List of 1-based page numbers where fitz detects images but MinerU
        extracted none.  Empty list if no gaps found or fitz unavailable.
    """
    if not _HAS_FITZ:
        _warn("F16: fitz (PyMuPDF) not available — cannot detect "
              "extraction gaps.")
        return []

    try:
        doc = fitz.open(str(pdf_path))
    except Exception as e:
        _warn(f"F16: could not open PDF with fitz: {e}")
        return []

    # Build set of pages that MinerU already extracted images from.
    # Page numbers in manifest are 1-based (0 means unknown).
    mineru_pages_with_images = set()
    for img_entry in mineru_manifest_images:
        page_num = img_entry.get("page")
        if page_num and page_num > 0:
            mineru_pages_with_images.add(page_num)

    gap_pages = []
    for page_idx in range(len(doc)):
        page_num = page_idx + 1  # 1-based
        page = doc[page_idx]

        # Check if fitz detects embedded images on this page
        try:
            page_images = page.get_images(full=True)
        except Exception:
            page_images = []

        if not page_images:
            continue  # No images on this page per fitz either

        # Filter out tiny images (icons, bullets, etc.) — width AND height
        # must be >= 50px to count as a real image.
        significant_images = []
        for img_info in page_images:
            xref = img_info[0]
            try:
                base_image = doc.extract_image(xref)
                if base_image:
                    w = base_image.get("width", 0)
                    h = base_image.get("height", 0)
                    if w >= 50 and h >= 50:
                        significant_images.append(img_info)
            except Exception:
                # If we cannot extract, still count it as potentially
                # significant (conservative approach).
                significant_images.append(img_info)

        if not significant_images:
            continue  # Only tiny images on this page

        # Gap: fitz sees significant images but MinerU extracted nothing
        if page_num not in mineru_pages_with_images:
            gap_pages.append(page_num)

    doc.close()

    if gap_pages:
        print(f"  F16: Detected {len(gap_pages)} page(s) with extraction "
              f"gaps: {gap_pages[:20]}"
              + ("..." if len(gap_pages) > 20 else ""))
    else:
        print("  F16: No extraction gaps detected (MinerU coverage OK)")

    return gap_pages


def _extract_fitz_fallback_images(
    pdf_path: Path,
    gap_pages: list,
    images_dir: Path,
    short_name: str,
    existing_manifest_images: list,
) -> list:
    """Extract images from gap pages using fitz as fallback for MinerU.

    For each page in gap_pages, uses PyMuPDF to extract embedded raster
    images and saves them to the pipeline images directory.  Updates naming
    to avoid collisions with existing MinerU-extracted images.

    Args:
        pdf_path: Path to the original PDF file.
        gap_pages: List of 1-based page numbers to extract from.
        images_dir: Pipeline images directory (same as MinerU images).
        short_name: Pipeline short name for file naming convention.
        existing_manifest_images: Current manifest images list (to
                                  determine next figure number).

    Returns:
        List of new manifest image dicts for fallback-extracted images.
        Empty list if extraction fails or no images found.
    """
    if not _HAS_FITZ:
        _warn("F16: fitz not available for fallback extraction.")
        return []

    if not gap_pages:
        return []

    try:
        doc = fitz.open(str(pdf_path))
    except Exception as e:
        _warn(f"F16: could not open PDF for fallback extraction: {e}")
        return []

    images_dir.mkdir(parents=True, exist_ok=True)

    # Determine next figure number from existing manifest
    max_fig_num = 0
    for img_entry in existing_manifest_images:
        fig_num = img_entry.get("figure_num", 0)
        if fig_num > max_fig_num:
            max_fig_num = fig_num

    # Track xrefs already extracted to avoid duplicates across pages
    # (same image can be referenced on multiple pages).
    extracted_xrefs = set()
    # RC5: SHA-256 content dedup — two different xrefs can contain
    # identical image bytes (e.g. the same figure embedded twice).
    # Maps hash hex → filename of first saved copy.
    seen_content_hashes: dict[str, str] = {}
    new_manifest_images = []
    fallback_idx = 0

    for page_num in gap_pages:
        page_idx = page_num - 1
        if page_idx < 0 or page_idx >= len(doc):
            continue

        page = doc[page_idx]
        try:
            page_images = page.get_images(full=True)
        except Exception:
            continue

        for img_info in page_images:
            xref = img_info[0]

            # Skip already-extracted xrefs (cross-page dedup)
            if xref in extracted_xrefs:
                continue

            try:
                base_image = doc.extract_image(xref)
            except Exception:
                continue

            if not base_image:
                continue

            # Filter tiny images (consistent with gap detection: both dims >= 50)
            w = base_image.get("width", 0)
            h = base_image.get("height", 0)
            if not (w >= 50 and h >= 50):
                continue

            # RC12: Skip full-page captures.
            # Fitz sometimes extracts the entire page as a single raster
            # image (e.g. A4 at 300 DPI = 2480x3508px).  These are not
            # useful individual images — they are full-page rasterizations.
            # Guard: if the image covers > 90% of the page in BOTH
            # dimensions (accounting for render DPI), skip it.
            page_w_pts = page.rect.width   # page width in points
            page_h_pts = page.rect.height  # page height in points
            if page_w_pts > 0 and page_h_pts > 0:
                # Image DPI is unknown; compare at multiple common DPIs.
                # If the image is >= 90% of page size at ANY standard DPI
                # (72, 96, 150, 200, 300), it is a full-page capture.
                _is_full_page = False
                for _dpi in (72, 96, 150, 200, 300):
                    page_w_px = page_w_pts * _dpi / 72.0
                    page_h_px = page_h_pts * _dpi / 72.0
                    if (w >= 0.9 * page_w_px and h >= 0.9 * page_h_px):
                        _is_full_page = True
                        break
                if _is_full_page:
                    print(f"    F16: Skipping full-page capture: "
                          f"{w}x{h}px (page {page_num})")
                    extracted_xrefs.add(xref)
                    continue

            image_bytes = base_image.get("image")
            if not image_bytes:
                continue

            # RC5: SHA-256 content dedup — skip if identical bytes
            # were already saved under a different xref.
            content_hash = hashlib.sha256(image_bytes).hexdigest()
            if content_hash in seen_content_hashes:
                print(f"    F16: Skipping duplicate fitz image on "
                      f"page {page_num} (SHA-256 matches "
                      f"{seen_content_hashes[content_hash]})")
                extracted_xrefs.add(xref)
                continue
            seen_content_hashes[content_hash] = (
                f"xref={xref}, page={page_num}")

            # Determine file extension from fitz
            ext = base_image.get("ext", "png")
            if ext not in ("png", "jpg", "jpeg", "bmp", "gif", "tiff"):
                ext = "png"

            extracted_xrefs.add(xref)
            fallback_idx += 1
            fig_num = max_fig_num + fallback_idx

            # Naming convention: {short_name}-fitz-fig{N}-page{P}.{ext}
            new_name = (f"{short_name}-fitz-fig{fig_num:03d}"
                        f"-page{page_num:03d}.{ext}")
            new_path = images_dir / new_name

            # Write image
            try:
                new_path.write_bytes(image_bytes)
            except Exception as e:
                _warn(f"F16: could not write fallback image "
                      f"{new_name}: {e}")
                continue
            if _ensure_max_dimension(new_path):
                # Re-read dimensions after resize
                try:
                    from PIL import Image as _PILDim  # noqa: PLC0415
                    with _PILDim.open(new_path) as _resized:
                        w, h = _resized.size
                except Exception:
                    pass

            # Check for near-black
            is_black = _is_near_black(new_path)
            was_rerendered = False
            if is_black:
                print(f"    F16: Near-black fallback image on page "
                      f"{page_num}, re-rendering with pdftoppm")
                rerender_name = (f"{short_name}-fitz-fig{fig_num:03d}"
                                 f"-page{page_num:03d}.png")
                rerender_path = images_dir / rerender_name
                success = _rerender_page_pdftoppm(
                    pdf_path, page_num, rerender_path)
                if success:
                    # Remove the near-black original if different path
                    if rerender_path != new_path and new_path.exists():
                        new_path.unlink()
                    new_path = rerender_path
                    new_name = rerender_name
                    was_rerendered = True
                    # Re-check dimensions after re-render
                    try:
                        from PIL import Image as _PILImage  # noqa: PLC0415
                        with _PILImage.open(new_path) as _img:
                            w, h = _img.size
                    except ImportError:
                        # PIL unavailable: retain original fitz dimensions.
                        # Near-black detection (_is_near_black) already
                        # required PIL, so this branch is theoretically
                        # unreachable, but guard is kept for safety.
                        pass
                    except Exception:
                        pass

            new_manifest_images.append({
                "index": len(existing_manifest_images) + fallback_idx - 1,
                "figure_num": fig_num,
                "filename": new_name,
                "path": str(new_path),
                "file_path": str(new_path),
                "page": page_num,
                "width": w,
                "height": h,
                "type_guess": "unknown",
                "extraction_source": "fitz_fallback",
                "mineru_source": "fitz_fallback",
                "section_context": {"page": page_num},
                "detected_caption": None,
                "near_black_detected": is_black,
                "rerendered": was_rerendered,
            })

    doc.close()

    if new_manifest_images:
        print(f"  F16: Extracted {len(new_manifest_images)} fallback "
              f"image(s) from {len(gap_pages)} gap page(s)")
    else:
        print(f"  F16: No extractable images found on gap pages")

    return new_manifest_images


def _normalize_mineru_output(
    pdf_path: Path,
    mineru_output_dir: Path,
    pipeline_output_md: Path,
    pipeline_images_dir: Path,
    short_name: str,
    cross_val_flag_rate: float,
    page_count: int,
) -> bool:
    """Normalize MinerU output to pipeline-compatible format.

    Takes the raw MinerU output directory and produces:
    1. A YAML-frontmattered markdown file at pipeline_output_md
    2. Renamed images in pipeline_images_dir
    3. An image-manifest.json compatible with prepare-image-analysis.py
    4. Near-black image detection + pdftoppm re-render

    Args:
        pdf_path: Path to the original PDF file.
        mineru_output_dir: MinerU's raw output directory
            (e.g., <tmpdir>/<stem>/auto/).
        pipeline_output_md: Where the final .md should be written.
        pipeline_images_dir: Where pipeline images should go.
        short_name: Pipeline short name for image naming.
        cross_val_flag_rate: The failure rate that triggered fallback.
        page_count: Number of pages in the PDF.

    Returns:
        True if normalization succeeded, False otherwise.
    """
    # ── 1. Find the MinerU markdown file ──
    md_files = list(mineru_output_dir.rglob("*.md"))
    if not md_files:
        _warn("MinerU normalization: no .md file found in output")
        return False
    source_md = max(md_files, key=lambda f: f.stat().st_size)
    md_content = source_md.read_text(encoding="utf-8")

    # ── 2. Find the content_list.json for page mapping ──
    content_list_files = list(mineru_output_dir.rglob("*content_list.json"))
    content_list = []
    if content_list_files:
        try:
            content_list = json.loads(
                content_list_files[0].read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError) as e:
            _warn(f"Could not parse MinerU content_list.json: {e}")

    # Build image-to-page mapping from content_list
    # content_list entries with type="image" have img_path and page_idx
    # Table entries (type="table") also have img_path when table_enable=false
    # (MinerU renders tables as screenshots in tables/ subdirectory)
    img_to_page = {}
    img_to_type = {}  # basename -> "image" or "table"
    for entry in content_list:
        entry_type = entry.get("type")
        if entry_type in ("image", "table") and "img_path" in entry:
            img_path_str = entry["img_path"]
            if not img_path_str:
                continue
            # img_path is relative, e.g. "images/hash.jpg" or "tables/hash.jpg"
            img_basename = Path(img_path_str).name
            # page_idx is 0-based
            img_to_page[img_basename] = entry.get("page_idx", -1)
            img_to_type[img_basename] = entry_type

    # Build page-to-heading mapping from content_list for section_context.
    # MinerU content_list entries with type="text" can be headings in two ways:
    #   1. text_level field present (1-4) — MinerU's standard format for titles
    #      (BlockType.Title → type="text" + text_level=N, raw text without "#")
    #   2. text starts with "#" — fallback for older MinerU versions that may
    #      embed markdown heading syntax in the text field
    # We collect all headings with their page_idx so we can find the nearest
    # heading above each image's page.
    _mineru_headings = []  # list of (page_idx, heading_text)
    for entry in content_list:
        if entry.get("type") == "text":
            text = entry.get("text", "").strip()
            text_level = entry.get("text_level")
            if text_level is not None and text_level >= 1:
                # MinerU title entry: raw text with text_level field
                if text:
                    _mineru_headings.append(
                        (entry.get("page_idx", -1), text))
            elif text.startswith("#"):
                # Fallback: markdown heading syntax in text field
                heading_text = re.sub(r'^#+\s*', '', text).strip()
                if heading_text:
                    _mineru_headings.append(
                        (entry.get("page_idx", -1), heading_text))

    def _find_mineru_section_heading(page_idx: int) -> str:
        """Find the nearest heading at or before the given page_idx."""
        best = ""
        for h_page, h_text in _mineru_headings:
            if h_page <= page_idx:
                best = h_text
            else:
                break  # content_list is in document order
        return best

    # ── 3. Collect MinerU images (from both images/ and tables/) ──
    mineru_images_dir = mineru_output_dir / "images"
    mineru_tables_dir = mineru_output_dir / "tables"
    _img_extensions = ("*.png", "*.jpg", "*.jpeg",
                       "*.svg", "*.bmp", "*.gif", "*.tiff", "*.tif")
    mineru_images = []
    if mineru_images_dir.exists():
        for ext in _img_extensions:
            mineru_images.extend(mineru_images_dir.glob(ext))
    # Also collect table screenshots (MinerU saves table regions to tables/)
    # These are rendered when table_enable=false (tables shown as images)
    mineru_table_images = []
    if mineru_tables_dir.exists():
        for ext in _img_extensions:
            mineru_table_images.extend(mineru_tables_dir.glob(ext))
    if mineru_table_images:
        print(f"    Found {len(mineru_table_images)} table screenshot(s) "
              f"in tables/ directory")
    # Combine: regular images first, then table images
    all_mineru_images = sorted(mineru_images, key=lambda p: p.name)
    all_mineru_images.extend(sorted(mineru_table_images, key=lambda p: p.name))
    # Track which source directory each image came from
    _table_image_names = {p.name for p in mineru_table_images}

    # ── 4. Prepare pipeline images directory ──
    pipeline_images_dir.mkdir(parents=True, exist_ok=True)

    # ── 5. Copy and rename images, detect near-black ──
    image_mapping = {}  # old_relative_path -> new_filename
    manifest_images = []
    near_black_count = 0
    rerendered_count = 0

    table_image_count = 0
    for idx, img_path in enumerate(all_mineru_images):
        old_basename = img_path.name
        is_table_img = old_basename in _table_image_names
        page_idx = img_to_page.get(old_basename, -1)
        page_num = page_idx + 1 if page_idx >= 0 else 0

        # Pipeline naming convention: short-name-fig<N>-page<P>.<ext>
        # Table screenshots get a "table_" prefix to distinguish them
        ext = img_path.suffix.lower()
        prefix = "table_" if is_table_img else ""
        new_name = (f"{short_name}-{prefix}fig{idx + 1:03d}"
                    f"-page{page_num:03d}{ext}")
        new_path = pipeline_images_dir / new_name
        if is_table_img:
            table_image_count += 1

        # Check for near-black before copying
        is_black = _is_near_black(img_path)
        was_rerendered = False
        if is_black:
            near_black_count += 1
            print(f"    Near-black detected: {old_basename} "
                  f"(page {page_num})")
            # Try to re-render with pdftoppm
            if page_num > 0:
                if ext != ".png":
                    new_name = new_name.rsplit(".", 1)[0] + ".png"
                    new_path = pipeline_images_dir / new_name
                success = _rerender_page_pdftoppm(
                    pdf_path, page_num, new_path)
                if success:
                    rerendered_count += 1
                    was_rerendered = True
                    print(f"    Re-rendered page {page_num} with pdftoppm")
                else:
                    # Fallback: copy the near-black image anyway
                    shutil.copy2(img_path, new_path)
                    _ensure_max_dimension(new_path)
                    print(f"    WARN: Could not re-render page {page_num}. "
                          "Keeping near-black image.")
            else:
                # No page info: copy as-is
                shutil.copy2(img_path, new_path)
                _ensure_max_dimension(new_path)
                print(f"    WARN: No page info for near-black image. "
                      "Keeping as-is.")
        else:
            shutil.copy2(img_path, new_path)
            _ensure_max_dimension(new_path)

        # Map old MinerU relative path to new filename
        # MinerU markdown uses: ![](images/hash.jpg) or ![](tables/hash.jpg)
        source_subdir = "tables" if is_table_img else "images"
        old_relative = f"{source_subdir}/{old_basename}"
        image_mapping[old_relative] = new_name

        # Build manifest entry
        # Get image dimensions if possible
        width, height = 0, 0
        if new_path.exists():
            try:
                from PIL import Image
                with Image.open(new_path) as img:
                    width, height = img.size
            except Exception:
                pass

        # Extract caption from content_list if available
        # Image entries use img_caption; table entries use table_caption
        caption = None
        expected_type = "table" if is_table_img else "image"
        for entry in content_list:
            if (entry.get("type") == expected_type
                    and Path(entry.get("img_path", "")).name == old_basename):
                caption_key = ("table_caption" if expected_type == "table"
                               else "img_caption")
                captions = entry.get(caption_key, [])
                if captions:
                    caption = " ".join(
                        c if isinstance(c, str)
                        else c.get("text", "")
                        for c in captions
                    ).strip() or None
                break

        manifest_images.append({
            "index": idx,
            "figure_num": idx + 1,
            "filename": new_name,
            "path": str(new_path),
            "file_path": str(new_path),
            "page": page_num if page_num > 0 else None,
            "width": width,
            "height": height,
            "type_guess": "table_screenshot" if is_table_img else "unknown",
            "extraction_source": "mineru",
            "mineru_source": "tables" if is_table_img else "images",
            "section_context": (
                {"heading": _find_mineru_section_heading(page_idx),
                 "page": page_num}
                if page_num > 0 else {}
            ),
            "detected_caption": caption,
            "near_black_detected": is_black,
            "rerendered": was_rerendered,
        })

    # ── 5b. F16: Hybrid fitz+MinerU fallback for missed images ──
    # Detect pages where fitz sees embedded images but MinerU extracted
    # none.  For those gap pages, extract images using fitz as fallback.
    _fitz_fallback_count = 0
    _fitz_fallback_images = []
    if _HAS_FITZ:
        _gap_pages = _detect_extraction_gaps(
            pdf_path=pdf_path,
            mineru_images_dir=pipeline_images_dir,
            mineru_manifest_images=manifest_images,
        )
        if _gap_pages:
            _fitz_fallback_images = _extract_fitz_fallback_images(
                pdf_path=pdf_path,
                gap_pages=_gap_pages,
                images_dir=pipeline_images_dir,
                short_name=short_name,
                existing_manifest_images=manifest_images,
            )
            if _fitz_fallback_images:
                _fitz_fallback_count = len(_fitz_fallback_images)
                # Extend the manifest with fallback images
                manifest_images.extend(_fitz_fallback_images)
                # Add markdown references for fallback images at end
                # of document (before YAML rewrite).  Each fallback
                # image gets a comment + image reference so it appears
                # in the converted markdown.
                _fb_md_lines = [
                    "\n\n<!-- F16: fitz fallback images for pages "
                    "missed by MinerU -->\n"
                ]
                for _fb_img in _fitz_fallback_images:
                    _fb_rel = (
                        f"{pipeline_output_md.stem}_images/"
                        f"{_fb_img['filename']}")
                    _fb_page = _fb_img.get("page", "?")
                    _fb_md_lines.append(
                        f"<!-- IMAGE: {_fb_img['filename']} "
                        f"(extracted via fitz fallback, page "
                        f"{_fb_page}) -->\n"
                        f"![fitz fallback page {_fb_page}]"
                        f"({_fb_rel})\n"
                    )
                md_content += "\n".join(_fb_md_lines)
                print(f"  F16: Added {_fitz_fallback_count} fallback "
                      f"image reference(s) to markdown")
    else:
        print("  F16: Skipped (fitz not available)")

    # ── 6. Rewrite image paths in markdown ──
    for old_path, new_name in image_mapping.items():
        # MinerU uses: ![](images/hash.jpg) or ![caption](images/hash.jpg)
        # Replace with pipeline-relative path using {stem}_images/ suffix
        # so links resolve correctly after --target-dir move places images
        # at {stem}_images/ alongside the .md file.
        pipeline_rel = f"{pipeline_output_md.stem}_images/{new_name}"
        md_content = md_content.replace(f"]({old_path})", f"]({pipeline_rel})")

    # ── 7. Inject YAML frontmatter ──
    _mineru_ver = _get_mineru_version()

    # Extract document title from MinerU markdown for the YAML header.
    # SYNC: this list must match INSTITUTIONAL_HEADERS in convert-paper.py
    _INSTITUTIONAL_HEADERS = [
        "health technology assessment", "statens legemiddelverk",
        "folkehelseinstituttet", "table of contents", "contents",
        "systematic review", "rapid review", "technology appraisal",
        "evidence report", "clinical practice guideline",
        "erasmus school of health policy", "university of oslo", "uio",
    ]
    _mineru_title = "Unknown"
    for _line in md_content.split('\n')[:80]:
        if _line.startswith('# ') and not _line.startswith('## '):
            _h1_text = re.sub(r'^#+\s*', '', _line).strip()
            if (len(_h1_text) > 5
                    and not any(ih in _h1_text.lower()
                                for ih in _INSTITUTIONAL_HEADERS)):
                _mineru_title = _h1_text
                break

    _safe_title = _mineru_title.replace('"', '\\"')

    # F14: Detect document domain from MinerU markdown content
    _domain, _domain_count, _domain_kws = _detect_document_domain(
        md_content)
    print(f"  Domain detection: {_domain} "
          f"({_domain_count} keyword(s): "
          f"{', '.join(_domain_kws[:5])})")

    # Total image count includes both MinerU and fitz fallback images
    _total_image_count = len(all_mineru_images) + _fitz_fallback_count

    yaml_header = (
        "---\n"
        f"source_file: {pdf_path.name}\n"
        f'title: "{_safe_title}"\n'
        f"document_domain: {_domain}\n"
        f"conversion_tool: magic-pdf (MinerU) v{_mineru_ver}\n"
        f"conversion_date: {datetime.now(timezone.utc).isoformat()}\n"
        f"fidelity_standard: best-effort (MinerU auto-fallback)\n"
        f"document_type: pdf\n"
        f"source_format: pdf\n"
        f"fallback_reason: cross-validation flag_rate "
        f"{cross_val_flag_rate:.1%} exceeded "
        f"{MINERU_FALLBACK_THRESHOLD:.0%} threshold\n"
        f"original_tool: pymupdf4llm (failed cross-validation)\n"
        f"pages: {page_count}\n"
        f"image_count: {_total_image_count}\n"
        f"regular_images: {len(all_mineru_images) - table_image_count}\n"
        f"table_screenshots: {table_image_count}\n"
        f"fitz_fallback_images: {_fitz_fallback_count}\n"
        f"near_black_images: {near_black_count}\n"
        f"rerendered_images: {rerendered_count}\n"
        f"image_notes: {'pending' if _total_image_count > 0 else 'none'}\n"
        "---\n\n"
    )

    # Remove any existing frontmatter from MinerU output (unlikely but safe)
    if md_content.startswith("---"):
        end_idx = md_content.find("---", 3)
        if end_idx != -1:
            md_content = md_content[end_idx + 3:].lstrip("\n")

    final_content = yaml_header + md_content

    # ── 8. Write the normalized markdown ──
    pipeline_output_md.parent.mkdir(parents=True, exist_ok=True)
    pipeline_output_md.write_text(final_content, encoding="utf-8")

    # ── 8b. Fix 3.6/M1: Classify MinerU images (remove bypass) ──
    # MinerU-extracted images previously skipped the classification chain
    # that pymupdf4llm images go through.  Run _classify_single_image()
    # on each manifest entry so decorative images are properly flagged.
    _mineru_dec_count = 0
    for _mi in manifest_images:
        # Build a minimal page_data dict for the classifier
        _mi_page_data = {
            "page": _mi.get("page", 0),
            "context": (_mi.get("section_context", {}).get("heading", "")
                        + " " + (_mi.get("detected_caption") or "")),
            "full_text": (_mi.get("section_context", {}).get("heading", "")
                          + " " + (_mi.get("detected_caption") or "")),
        }
        # xref_counts not available for MinerU images; pass empty dict
        _classify_single_image(
            _mi, _mi_page_data, xref_counts={}, total_pages=page_count)
        if not _mi.get("is_substantive", True):
            _mineru_dec_count += 1
    if _mineru_dec_count > 0:
        print(f"  Fix 3.6: Classified {_mineru_dec_count} MinerU image(s) "
              f"as decorative")

    # ── 9. Write the image manifest ──
    manifest_data = {
        "source": "mineru",
        "source_version": f"magic-pdf {_mineru_ver}",
        "md_file": str(pipeline_output_md),
        "images_dir": str(pipeline_images_dir),
        "image_count": len(manifest_images),
        "table_image_count": table_image_count,
        "fitz_fallback_count": _fitz_fallback_count,
        "near_black_count": near_black_count,
        "rerendered_count": rerendered_count,
        "generated": datetime.now(timezone.utc).isoformat(),
        "images": manifest_images,
    }
    manifest_path = pipeline_images_dir / "image-manifest.json"
    manifest_path.write_text(json.dumps(manifest_data, indent=2),
                             encoding="utf-8")

    print(f"  MinerU normalization complete:")
    print(f"    Markdown: {pipeline_output_md} "
          f"({len(final_content):,} bytes)")
    print(f"    Images:   {len(all_mineru_images)} MinerU + "
          f"{_fitz_fallback_count} fitz fallback = "
          f"{len(manifest_images)} total → {pipeline_images_dir}")
    if table_image_count > 0:
        print(f"    (includes {table_image_count} table screenshot(s) "
              f"from tables/ directory)")
    if _fitz_fallback_count > 0:
        print(f"    Fitz fallback: {_fitz_fallback_count} image(s) "
              f"recovered from gap pages")
    print(f"    Manifest: {manifest_path}")
    if near_black_count > 0:
        print(f"    Near-black: {near_black_count} detected, "
              f"{rerendered_count} re-rendered with pdftoppm")

    return True


def _trigger_mineru_fallback(
    input_file: Path,
    output_md: Path,
    images_dir: Path,
    short_name: str,
    checkpoint: dict,
    checkpoint_path: Path,
    cross_val_flag_rate: float,
    page_count: int,
) -> bool:
    """Discard pymupdf4llm output and retry with MinerU.

    Called when cross-validation failure rate exceeds
    MINERU_FALLBACK_THRESHOLD.  Runs MinerU in a temp directory,
    normalizes the output, and replaces the pymupdf4llm artifacts.

    Note: target_dir handling is done by the caller, which resolves
    output_md and images_dir to the correct paths before calling.

    Args:
        input_file: Path to the source PDF.
        output_md: Path for the pipeline output markdown.
        images_dir: Path for the pipeline images directory.
        short_name: Pipeline short name.
        checkpoint: Current pipeline checkpoint dict.
        checkpoint_path: Path to the checkpoint JSON.
        cross_val_flag_rate: The failure rate that triggered this.
        page_count: Number of pages in the PDF.

    Returns:
        True if MinerU fallback succeeded, False otherwise.
    """
    if not _mineru_available():
        _warn("MinerU fallback requested but MinerU venv not found.")
        return False

    print(f"\n{'─' * 40}")
    print("Step 1b-FALLBACK: MinerU Cross-Validation Fallback")
    print('─' * 40)
    print(f"  Flag rate: {cross_val_flag_rate:.1%} >= "
          f"{MINERU_FALLBACK_THRESHOLD:.0%} threshold")
    print(f"  Pages: {page_count}")
    print("  Auto-switching to MinerU extractor...")

    # ── Run MinerU in a temp directory ──
    # MinerU creates: <output_dir>/<pdf_stem>/auto/<pdf_stem>.md
    with tempfile.TemporaryDirectory(prefix="mineru-fallback-") as tmpdir:
        tmpdir_path = Path(tmpdir)

        # Build MinerU command with CPU-safety
        env = os.environ.copy()
        env["CUDA_VISIBLE_DEVICES"] = ""
        env["PYTORCH_MPS_DISABLE"] = "1"
        env["PYTORCH_ENABLE_MPS_FALLBACK"] = "0"

        # Use the MinerU venv's magic-pdf CLI
        magic_pdf_bin = MINERU_VENV / "bin" / "magic-pdf"
        if magic_pdf_bin.exists():
            cmd = [
                str(magic_pdf_bin),
                "-p", str(input_file),
                "-o", str(tmpdir_path),
                "-m", "auto",
            ]
        else:
            cmd = [
                str(MINERU_PYTHON),
                "-m", "magic_pdf.cli.main",
                "-p", str(input_file),
                "-o", str(tmpdir_path),
                "-m", "auto",
            ]

        # Quote args with spaces for readable log output
        # (subprocess.run with list handles spaces correctly regardless)
        _quoted_cmd = [
            f'"{c}"' if " " in c else c for c in cmd
        ]
        print(f"  Command: {' '.join(_quoted_cmd)}")
        print("  Running MinerU (CPU-only, timeout 10 min)...")

        try:
            result = subprocess.run(
                cmd,
                env=env,
                capture_output=True,
                text=True,
                timeout=600,
            )
        except subprocess.TimeoutExpired:
            _warn("MinerU fallback timed out after 10 minutes. "
                  "Continuing with pymupdf4llm output.")
            checkpoint["mineru_fallback_timeout"] = True
            return False
        except Exception as e:
            _warn(f"MinerU fallback subprocess failed: {e}")
            return False

        if result.returncode != 0:
            _warn(f"MinerU fallback exited with code {result.returncode}")
            if result.stderr:
                _warn(f"  stderr: {result.stderr[:500]}")
            return False

        # ── Locate MinerU output directory ──
        # MinerU creates: <tmpdir>/<pdf_stem>/auto/
        pdf_stem = input_file.stem
        mineru_auto_dir = tmpdir_path / pdf_stem / "auto"
        if not mineru_auto_dir.exists():
            # Try finding the auto dir by scanning
            auto_dirs = list(tmpdir_path.rglob("auto"))
            if auto_dirs:
                mineru_auto_dir = auto_dirs[0]
            else:
                _warn("MinerU fallback: cannot find 'auto' output directory")
                return False

        print(f"  MinerU output found: {mineru_auto_dir}")

        # ── Back up pymupdf4llm artifacts (restore on failure) ──
        # BUG FIX (S61): Previously deleted pymupdf4llm output BEFORE
        # verifying MinerU normalization succeeded.  When MinerU failed,
        # both outputs were lost.  Now we rename to .bak, attempt MinerU,
        # and restore on failure.
        _backup_md = output_md.with_suffix(".md.pymupdf4llm_backup")
        _backup_images = (
            images_dir.with_name(images_dir.name + "_pymupdf4llm_backup")
            if images_dir.exists() else None
        )
        if output_md.exists():
            output_md.rename(_backup_md)
        if _backup_images and images_dir.exists():
            images_dir.rename(_backup_images)

        # ── Normalize MinerU output ──
        success = _normalize_mineru_output(
            pdf_path=input_file,
            mineru_output_dir=mineru_auto_dir,
            pipeline_output_md=output_md,
            pipeline_images_dir=images_dir,
            short_name=short_name,
            cross_val_flag_rate=cross_val_flag_rate,
            page_count=page_count,
        )

        if not success:
            # ── Restore pymupdf4llm output ──
            if _backup_md.exists():
                _backup_md.rename(output_md)
            if _backup_images and _backup_images.exists():
                _backup_images.rename(images_dir)
            _warn("MinerU normalization failed. "
                  "Restored pymupdf4llm output.")
            return False

        # ── MinerU succeeded — clean up backups ──
        if _backup_md.exists():
            _backup_md.unlink()
        if _backup_images and _backup_images.exists():
            shutil.rmtree(_backup_images, ignore_errors=True)

    # ── Update checkpoint ──
    checkpoint["extractor"] = "mineru"
    checkpoint["extractor_used"] = "mineru"
    checkpoint["mineru_fallback_triggered"] = True
    checkpoint["mineru_fallback_flag_rate"] = round(cross_val_flag_rate, 4)
    try:
        checkpoint_path.parent.mkdir(parents=True, exist_ok=True)
        checkpoint_path.write_text(json.dumps(checkpoint, indent=2))
    except Exception as e:
        _warn(f"Could not update checkpoint after MinerU fallback: {e}")

    print("  MinerU fallback: SUCCESS")
    return True


# ═══════════════════════════════════════════════════════════════════════════
# v3.1 ORGANIZATION HELPERS (R1, R2, R3, R4, R5, R14)
# ═══════════════════════════════════════════════════════════════════════════

def atomic_move(src: Path, dst: Path) -> None:
    """Move a file atomically. Cross-volume safe via copy-verify-delete.

    R1 helper — provided verbatim in PIPELINE-V31-REQUIREMENTS.md.
    On same filesystem: os.rename is atomic (APFS).
    On different filesystems: copy2 → SHA-256 verify → delete source.
    """
    dst.parent.mkdir(parents=True, exist_ok=True)
    try:
        src.rename(dst)  # atomic on same filesystem
    except OSError:
        # Cross-filesystem: copy then verify then delete
        shutil.copy2(src, dst)
        if _compute_sha256(src) != _compute_sha256(dst):
            dst.unlink()
            raise RuntimeError(f"Copy verification failed: {src}")
        src.unlink()


def verify_conversion_output(md_path: Path) -> tuple:
    """R5: Pre-deletion verification.

    Returns (passed: bool, warnings: list, errors: list, has_frontmatter: bool).

    Checks (in order):
      1. .md file exists                   → blocking
      2. File size > 0 bytes               → blocking
      3. At least one non-whitespace line   → blocking
      4. YAML frontmatter present (---)     → advisory warning only

    If any blocking check fails, returns (False, warnings, errors, False).
    If only advisory checks fail, returns (True, warnings, [], False).
    """
    warnings = []
    errors = []
    has_frontmatter = False

    # Check 1: file exists
    if not md_path.exists():
        errors.append(f"Output .md file does not exist: {md_path}")
        return (False, warnings, errors, has_frontmatter)

    # Check 2: file size > 0
    size = md_path.stat().st_size
    if size == 0:
        errors.append(f"Output .md file is empty (0 bytes): {md_path}")
        return (False, warnings, errors, has_frontmatter)

    # Check 3: at least one non-whitespace line
    try:
        text = md_path.read_text(encoding="utf-8")
        has_content = any(line.strip() for line in text.splitlines())
        if not has_content:
            errors.append(
                f"Output .md file contains only whitespace ({size} bytes): "
                f"{md_path}"
            )
            return (False, warnings, errors, has_frontmatter)
    except Exception as e:
        errors.append(f"Could not read output .md file: {e}")
        return (False, warnings, errors, has_frontmatter)

    # Check 4: YAML frontmatter (advisory only)
    first_line = text.splitlines()[0] if text.splitlines() else ""
    if first_line.strip().startswith("---"):
        has_frontmatter = True
    else:
        warnings.append(
            f"Output .md lacks YAML frontmatter (first line: "
            f"{first_line[:40]!r}): {md_path}"
        )

    return (True, warnings, errors, has_frontmatter)


def check_already_organized(source_file: Path, target_dir: Path) -> bool:
    """R14: Idempotency check. Returns True if this file is already organized.

    Checks whether _originals/ already contains a file with the same name
    AND matching SHA-256 hash as the source file. If so, the organize steps
    can be skipped.
    """
    originals_dir = target_dir / ORIGINALS_SUBDIR
    dest_file = originals_dir / source_file.name

    if not dest_file.exists():
        return False

    # File exists at destination — verify SHA-256 match
    src_hash = _compute_sha256(source_file)
    dst_hash = _compute_sha256(dest_file)
    return src_hash == dst_hash


def move_images_dir(source_images: Path, target_images: Path,
                    md_path: Path, dry_run: bool = False) -> list:
    """R3: Move image directory to target location. Update .md references.

    Returns list of (action, description) tuples for reporting.
    Handles:
    - file=None manifest entries (chart rendering failures) gracefully
    - Already-in-place images (no-op)
    - Image path reference updates in the .md file
    """
    actions = []

    if not source_images.exists():
        return actions

    if source_images.resolve() == target_images.resolve():
        actions.append(("SKIP", f"Images already at {target_images}"))
        return actions

    if dry_run:
        actions.append(("DRY RUN", f"Would move images {source_images} → {target_images}"))
        return actions

    # Move the directory tree
    target_images.parent.mkdir(parents=True, exist_ok=True)
    if target_images.exists():
        # Merge: copy contents into existing dir
        for item in source_images.iterdir():
            dest_item = target_images / item.name
            if item.is_file() and not dest_item.exists():
                shutil.copy2(item, dest_item)
            elif item.is_dir() and not dest_item.exists():
                shutil.copytree(item, dest_item)
        shutil.rmtree(source_images)
    else:
        shutil.move(str(source_images), str(target_images))
    actions.append(("PLACED", f"Images dir → {target_images}"))

    # Update image references in .md if paths changed
    if md_path.exists():
        try:
            content = md_path.read_text(encoding="utf-8")
            old_ref = source_images.name  # e.g. "My File_images"
            new_ref = target_images.name  # should be the same basename
            # Only rewrite if the relative path prefix changed
            # (i.e. images were in a different parent dir)
            if source_images.parent.resolve() != target_images.parent.resolve():
                # References are typically relative: "stem_images/file.png"
                # If .md moved too, references may still be valid.
                # Update any absolute references that point to old location.
                old_abs = str(source_images)
                new_abs = str(target_images)
                if old_abs in content:
                    content = content.replace(old_abs, new_abs)
                    md_path.write_text(content, encoding="utf-8")
                    actions.append(("UPDATED", "Image references in .md updated"))
        except Exception as e:
            actions.append(("WARNING", f"Could not update image refs in .md: {e}"))

    return actions


def _format_size(size_bytes: int) -> str:
    """R11: Human-readable file size (bytes → KB → MB)."""
    if size_bytes < 1024:
        return f"{size_bytes} bytes"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    else:
        return f"{size_bytes / (1024 * 1024):.1f} MB"


def _truncate_path(path_str: str, max_len: int = 60) -> str:
    """R11: Truncate long paths to last max_len chars with \u2026 prefix."""
    if len(path_str) <= max_len:
        return path_str
    return "\u2026" + path_str[-(max_len - 1):]


def _read_image_count_from_manifest(target_dir: Path,
                                     input_stem: str) -> tuple:
    """R7: Read image counts from the final manifest JSON.

    Returns (total_images, unique_images, duplicates_skipped).
    Chart-rendered PNGs are included because convert-office.py adds them
    to the manifest before writing it.
    """
    manifest_candidates = [
        target_dir / f"{input_stem}_manifest.json",
        target_dir / f"{input_stem.lower().replace(' ', '-')}_manifest.json",
    ]
    for candidate in manifest_candidates:
        if candidate.exists():
            try:
                data = json.loads(candidate.read_text(encoding="utf-8"))
                images = data.get("images", [])
                total = data.get("total_extracted",
                                data.get("total_images", len(images)))
                unique = data.get("image_count", len(images))
                dupes = data.get("duplicates_skipped",
                                 sum(1 for img in images
                                     if img.get("is_duplicate", False)))
                return (total, unique, dupes)
            except (json.JSONDecodeError, KeyError):
                pass
    return (0, 0, 0)


def check_registry_duplicate(source_hash: str,
                              target_dir: Optional[Path]) -> Optional[dict]:
    """R12: Check registry for existing conversion of this file.

    Returns the matching registry entry if found with same hash AND same
    target_dir. Returns None if no match or different target_dir.
    """
    if not REGISTRY_PATH.exists():
        return None

    try:
        with open(REGISTRY_PATH) as f:
            registry = json.load(f)
    except (json.JSONDecodeError, ValueError):
        return None

    for entry in registry.get("conversions", []):
        if entry.get("source_hash") != source_hash:
            continue
        # Same hash found — check target_dir match
        if target_dir is None:
            # No target-dir: any existing entry counts
            return entry
        entry_target = entry.get("target_dir", "")
        if entry_target and Path(entry_target).resolve() == target_dir.resolve():
            return entry
        # Same hash but different target_dir: not a duplicate
        # (user intentionally placing in new location)
    return None


def update_registry_organized(source_hash: str, source_file: Path,
                               output_md: Path, target_dir: Path,
                               extractor: str,
                               images_dir: Optional[Path] = None,
                               image_index_meta: Optional[dict] = None
                               ) -> None:
    """R10 + R21: Extend registry entry with organized paths and image index
    metadata after file organization.

    Writes a NEW entry with organized_* fields rather than modifying
    existing v3.0 entries (registry migration rule).

    R21: When image_index_meta is provided, adds 7 image-related fields:
        image_index_path, image_index_generated_at, total_pages,
        pages_with_images, total_images_detected, substantive_images,
        has_testable_images.
    Old entries without these fields are backward compatible (not modified).
    """
    REGISTRY_PATH.parent.mkdir(parents=True, exist_ok=True)

    originals_dir = target_dir / ORIGINALS_SUBDIR

    new_entry = {
        "source_hash": source_hash,
        "source_path": str(source_file.resolve()),
        "output_path": str(output_md.resolve()),
        "organized_source_path": str(originals_dir / source_file.name),
        "organized_output_path": str(output_md.resolve()),
        "organized_images_path": str(images_dir) if images_dir else "",
        "target_dir": str(target_dir),
        "extractor_used": extractor,
        "converted_at": datetime.now().isoformat(),
        "organized_at": datetime.now().isoformat(),
        "pipeline_version": PIPELINE_VERSION,
    }

    # R21: Add image index metadata if available
    if image_index_meta is not None:
        new_entry["image_index_path"] = image_index_meta.get(
            "image_index_path", "")
        new_entry["image_index_generated_at"] = image_index_meta.get(
            "image_index_generated_at", "")
        new_entry["total_pages"] = image_index_meta.get(
            "total_pages", 0)
        new_entry["pages_with_images"] = image_index_meta.get(
            "pages_with_images", 0)
        new_entry["total_images_detected"] = image_index_meta.get(
            "total_images_detected", 0)
        new_entry["substantive_images"] = image_index_meta.get(
            "substantive_images", 0)
        new_entry["has_testable_images"] = image_index_meta.get(
            "has_testable_images", False)

    # Lock the entire read-modify-write cycle to prevent concurrent
    # lost updates (MINOR-2 fix).
    # NOTE: The .json.lock file is intentionally persistent (never deleted).
    # fcntl.flock() works on existing files, and keeping the file avoids a
    # race condition where two processes try to create it simultaneously.
    _lock_path = REGISTRY_PATH.with_suffix(".json.lock")
    with open(_lock_path, "w") as _lf:
        fcntl.flock(_lf, fcntl.LOCK_EX)
        try:
            registry = {"conversions": []}
            if REGISTRY_PATH.exists():
                try:
                    with open(REGISTRY_PATH) as f:
                        registry = json.load(f)
                    if not isinstance(registry, dict) or "conversions" not in registry:
                        raise ValueError("malformed registry")
                except (json.JSONDecodeError, ValueError) as e:
                    _warn(f"Registry corrupt ({e}). Using empty registry for R10 update.")
                    registry = {"conversions": []}

            # Remove existing entry with same hash AND same target_dir (dedup)
            # Preserve v3.0 entries that lack organized_* fields
            kept = []
            for entry in registry.get("conversions", []):
                if entry.get("source_hash") == source_hash:
                    entry_target = entry.get("target_dir", "")
                    if entry_target and Path(entry_target).resolve() == target_dir.resolve():
                        continue  # Replace this entry
                    if not entry_target and "organized_source_path" not in entry:
                        # v3.0 entry without target_dir: preserve it
                        kept.append(entry)
                        continue
                kept.append(entry)

            kept.append(new_entry)
            registry["conversions"] = kept

            # Atomic write: temp file then replace
            tmp_path = REGISTRY_PATH.with_suffix(".json.tmp")
            with open(tmp_path, "w") as f:
                json.dump(registry, f, indent=2)
            tmp_path.replace(REGISTRY_PATH)
        finally:
            fcntl.flock(_lf, fcntl.LOCK_UN)

    print(f"\n  Registry updated (R10): {REGISTRY_PATH}")
    print(f"    organized_source_path: {new_entry['organized_source_path']}")
    print(f"    organized_output_path: {new_entry['organized_output_path']}")
    print(f"    organized_at: {new_entry['organized_at']}")
    if image_index_meta:
        print(f"    image_index_path: "
              f"{new_entry.get('image_index_path', 'N/A')}")
        print(f"    has_testable_images: "
              f"{new_entry.get('has_testable_images', False)}")


def append_issue_log(target_dir: Path, source_file: Path,
                     output_md: Optional[Path], extractor: str,
                     issue_type: str, severity: str,
                     details: str, action_taken: str) -> None:
    """R6: Append a structured issue entry to CONVERSION-ISSUES.md.

    Only called when --target-dir is provided (R16 compatibility).
    Creates the file on first issue. Never overwrites existing entries.
    """
    log_path = target_dir / "CONVERSION-ISSUES.md"
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    source_name = source_file.name

    entry = (
        f"\n## [{timestamp}] {source_name} — {severity}\n\n"
        f"- **Source:** {source_file.resolve()}\n"
        f"- **Output:** {output_md.resolve() if output_md else 'N/A'}\n"
        f"- **Pipeline version:** v{PIPELINE_VERSION}\n"
        f"- **Extractor used:** {extractor}\n"
        f"- **Issue type:** {issue_type}\n"
        f"- **Details:** {details}\n"
        f"- **Action taken:** {action_taken}\n"
    )

    try:
        if log_path.exists():
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(entry)
        else:
            with open(log_path, "w", encoding="utf-8") as f:
                f.write("# Conversion Issues Log\n\n")
                f.write("Pipeline-generated issue tracking. "
                        "Entries are appended automatically.\n")
                f.write(entry)
    except Exception as e:
        _warn(f"Could not write issue log: {e}")


def generate_visual_report(source_file: Path, target_dir: Optional[Path],
                            output_md: Path, source_hash: str,
                            extractor: str, input_stem: str,
                            move_actions: list, place_actions: list,
                            cleanup_actions: list, issues: list,
                            verification_passed: bool,
                            has_frontmatter: bool,
                            dry_run: bool = False,
                            status: str = "COMPLETE",
                            image_index_meta: Optional[dict] = None
                            ) -> str:
    """R7 + R11 + R19: Generate the visual conversion report.

    Returns the report as a string. Caller prints to stdout and
    optionally writes to disk.

    R19 extension: When image_index_meta is provided, adds an
    IMAGE INDEX section to the report.
    """
    now = datetime.now()
    timestamp_display = now.strftime("%Y-%m-%d %H:%M:%S")
    timestamp_file = now.strftime("%Y%m%d-%H%M%S")

    # Read image count from final manifest (R7 critical constraint)
    total_img, unique_img, dupes_img = (0, 0, 0)
    if target_dir:
        total_img, unique_img, dupes_img = \
            _read_image_count_from_manifest(target_dir, input_stem)
    else:
        total_img, unique_img, dupes_img = \
            _read_image_count_from_manifest(output_md.parent, input_stem)

    # Build report
    lines = []
    header_source = _truncate_path(str(source_file))
    w = 62  # inner width of the box

    lines.append("╔" + "═" * w + "╗")
    lines.append("║  PIPELINE v3.1 — FILE ORGANIZATION REPORT"
                 + " " * (w - 42) + "║")
    lines.append(f"║  Source: {header_source}"
                 + " " * max(0, w - 10 - len(header_source)) + "║")
    lines.append(f"║  Completed: {timestamp_display}"
                 + " " * max(0, w - 13 - len(timestamp_display)) + "║")
    lines.append("╚" + "═" * w + "╝")
    lines.append("")

    # CONVERSIONS section
    lines.append("CONVERSIONS")
    lines.append("─" * 39)
    md_size_str = ""
    if output_md.exists():
        md_size_str = f" ({_format_size(output_md.stat().st_size)})"
    lines.append(f"  ✓ {source_file.name}")
    lines.append(f"        → {input_stem}.md{md_size_str}")
    lines.append(f"        Extractor: {extractor}")
    if total_img > 0:
        img_detail = f"{total_img} extracted"
        if unique_img != total_img or dupes_img > 0:
            img_detail += f", {unique_img} unique ({dupes_img} duplicates skipped)"
        lines.append(f"        Images:    {img_detail}")
    lines.append("")

    # IMAGE INDEX section (R19)
    if image_index_meta:
        lines.append("IMAGE INDEX")
        lines.append("─" * 39)
        idx_path = image_index_meta.get("image_index_path", "")
        idx_name = Path(idx_path).name if idx_path else "N/A"
        lines.append(f"  GENERATED  {idx_name}")
        lines.append(f"             Pages scanned: "
                     f"{image_index_meta.get('total_pages', 0)}")
        lines.append(f"             Pages with images: "
                     f"{image_index_meta.get('pages_with_images', 0)}")
        lines.append(f"             Total images: "
                     f"{image_index_meta.get('total_images_detected', 0)}")
        total_imgs = image_index_meta.get("total_images_detected", 0)
        subst_imgs = image_index_meta.get("substantive_images", 0)
        deco_imgs = total_imgs - subst_imgs
        subst_pct = (subst_imgs / total_imgs * 100) if total_imgs > 0 else 0
        deco_pct = (deco_imgs / total_imgs * 100) if total_imgs > 0 else 0
        lines.append(f"             Substantive: "
                     f"{subst_imgs} ({subst_pct:.0f}%)")
        lines.append(f"             Decorative filtered: "
                     f"{deco_imgs} ({deco_pct:.0f}%)")
        lines.append("")

    # FILE MOVEMENTS section
    lines.append("FILE MOVEMENTS")
    lines.append("─" * 39)
    if target_dir:
        dr = " (DRY RUN)" if dry_run else ""
        for action_type, desc in move_actions:
            lines.append(f"  {action_type}{dr}  {desc}")
        for action_type, desc in place_actions:
            lines.append(f"  {action_type}{dr}  {desc}")
        if not move_actions and not place_actions:
            lines.append("  No file movements performed.")
    else:
        lines.append("  No target-dir, organization skipped.")
    lines.append("")

    # DELETED section
    lines.append("DELETED (intermediate files)")
    lines.append("─" * 39)
    if cleanup_actions:
        for action_type, desc in cleanup_actions:
            lines.append(f"  {action_type}  {_truncate_path(desc)}")
    else:
        lines.append("  No intermediate files to clean up.")
    lines.append("")

    # ISSUES section
    lines.append("ISSUES")
    lines.append("─" * 39)
    if issues:
        for issue in issues:
            icon = "⚠" if issue.get("severity") != "CRITICAL" else "✗"
            lines.append(f"  {icon}  {issue['details']}")
        if target_dir:
            lines.append(f"     Details: "
                         f"{_truncate_path(str(target_dir / 'CONVERSION-ISSUES.md'))}")
    else:
        lines.append("  No issues.")
    lines.append("")

    # VERIFICATION section
    lines.append("VERIFICATION")
    lines.append("─" * 39)
    if verification_passed:
        if output_md.exists():
            sz = _format_size(output_md.stat().st_size)
            lines.append(f"  ✓ Output .md exists and non-empty ({sz})")
        else:
            lines.append("  ✓ Output .md verified")
        if has_frontmatter:
            lines.append("  ✓ YAML frontmatter present")
        else:
            lines.append("  ⚠ YAML frontmatter missing (advisory)")
        if dry_run:
            lines.append(f"  ○ Registry would be updated (SHA-256: {source_hash.removeprefix('sha256:')[:20]}...)")
        else:
            lines.append(f"  ✓ Registry updated (SHA-256: {source_hash.removeprefix('sha256:')[:20]}...)")
    else:
        lines.append("  ✗ Verification failed")
    lines.append("")

    # STATUS line
    lines.append(f"STATUS: {status}")
    lines.append("╚" + "═" * w + "╝")

    return "\n".join(lines)


def cleanup_intermediate_files(output_md: Path, source_file: Path,
                               checkpoint_path: Path,
                               dry_run: bool = False) -> list:
    """R4: Clean up intermediate/temp files after successful organization.

    Returns list of (action, path) tuples for reporting.

    CRITICAL: Does NOT glob /tmp/soffice-* — chart rendering cleans its
    own temp dirs in convert-office.py's finally block.
    Only cleans files THIS pipeline code created.
    """
    actions = []
    targets = []

    # 1. Pipeline checkpoint file
    if checkpoint_path.exists():
        targets.append(checkpoint_path)

    # 2. Intermediate .txt sidecar files (produced by some extractors)
    txt_sidecar = output_md.with_suffix(".txt")
    if txt_sidecar.exists() and txt_sidecar != source_file:
        targets.append(txt_sidecar)

    # 3. MinerU working directories (images/mineru/ when final images
    #    have been moved to _images/)
    mineru_work_dir = output_md.parent / "images" / "mineru"
    if mineru_work_dir.exists():
        targets.append(mineru_work_dir)

    # 4. Empty images/ parent dir if mineru was the only child
    images_parent = output_md.parent / "images"
    if images_parent.exists() and not any(images_parent.iterdir()):
        targets.append(images_parent)

    for target in targets:
        if dry_run:
            actions.append(("DRY RUN", f"Would delete {target}"))
        else:
            try:
                if target.is_dir():
                    shutil.rmtree(target)
                else:
                    target.unlink()
                actions.append(("DELETED", str(target)))
            except Exception as e:
                actions.append(("WARNING", f"Could not delete {target}: {e}"))

    return actions


# ═══════════════════════════════════════════════════════════════════════════
# R19: IMAGE INDEX GENERATION (Per-File)
# ═══════════════════════════════════════════════════════════════════════════


def _is_blank_image(img_path: str,
                    std_threshold: float = 5.0,
                    size_threshold: int = 2000) -> bool:
    """FIX-2 + M1/M2: Detect blank, near-blank, or decorative images.

    An image is blank/decorative if:
      - Its file size is below size_threshold bytes (tiny placeholder).
      - Its pixel standard deviation is below std_threshold (near-
        uniform color, e.g. all-white).
      - M1 heuristic: very few unique colors (< 32 in RGB) indicating
        a solid color block, gradient bar, or simple decorative element.
        Conservative: only triggers when both unique colors are low AND
        image is not predominantly white (mean < 240), to avoid
        flagging simple diagrams on white backgrounds.
      - M2 heuristic: near-uniform dark image (std < 10, mean < 50,
        unique colors < 16) indicating a decorative cover/background
        pattern such as dot grids on dark pages.

    Returns True if the image is blank/near-blank/decorative.
    """
    file_size = 0
    try:
        import os as _os
        file_size = _os.path.getsize(img_path)
        if file_size < size_threshold:
            return True
    except OSError:
        pass

    try:
        from PIL import Image as _PILImage
        import numpy as _np
        img = _PILImage.open(img_path)
        # M-1 FIX: Detect alpha channel but do NOT early-return.
        # The old code returned False for any RGBA image with alpha < 250,
        # which bypassed ALL M1/M2 decorative heuristics.  Decorative
        # images (CC badges, logos, color-block overlays) can have alpha
        # channels too.  Now: we note the presence of transparency and
        # only use it to guard the original "near-uniform gray" blank
        # check (which would false-positive on transparent images).
        _has_significant_alpha = False
        if img.mode == 'RGBA':
            alpha = _np.array(img.split()[-1])
            if alpha.mean() < 250:  # Has significant transparency
                _has_significant_alpha = True

        # Compute grayscale stats on RGB channels (ignoring alpha) so
        # that M1/M2 heuristics work correctly on RGBA images.
        gray = img.convert("L")  # grayscale (drops alpha)
        arr = _np.array(gray)
        gray_std = arr.std()
        gray_mean = arr.mean()

        # Tier 2: Near-black composite (catches anti-aliased blank fragments)
        # Anti-aliased edges on near-black blanks push std > 5, bypassing
        # the original std check.  This OR branch catches them when:
        #   file_size < 5KB AND mean pixel < 30 AND < 16 unique colors.
        # Compute unique colors early for this check (reused later by M1).
        unique_colors = -1  # sentinel: not computed
        if img.mode in ('RGB', 'RGBA'):
            rgb_data = img.convert('RGB')
            # Sample up to 50000 pixels for performance on large images
            w_px, h_px = rgb_data.size
            if w_px * h_px > 50000:
                rgb_data = rgb_data.resize(
                    (min(w_px, 250), min(h_px, 200)),
                    _PILImage.NEAREST)
            unique_colors = len(set(rgb_data.getdata()))

        if (file_size > 0 and file_size < 5000
                and gray_mean < 30
                and unique_colors >= 0 and unique_colors < 16):
            return True

        # Tier 3 (original): near-uniform pixel values.
        # GUARDED by alpha: transparent PNGs with uniform background
        # would be a false positive here (the "uniform" color is the
        # composited-on-white result, not the actual image content).
        if gray_std < std_threshold and not _has_significant_alpha:
            return True

        # M1/M2 heuristics: color-block and dark-cover detection.
        # These run regardless of alpha channel — decorative images
        # can be transparent PNGs.
        # unique_colors already computed above (Tier 2 block).
        # M1 heuristic: color-block detection.
        # Solid color blocks and gradient bars used as journal styling
        # have very few unique colors (< 32).  We require the image
        # to NOT be mostly white (mean < 240) to avoid false positives
        # on simple line-art diagrams rendered on white backgrounds.
        # Also require gray_std < 30 to avoid flagging images with
        # genuine content that happen to use few colors (e.g. simple
        # charts with 10 distinct colored bars, two-tone diagrams).
        if (unique_colors >= 0 and unique_colors < 32
                and gray_mean < 240 and gray_std < 30):
            return True

        # M2 dark-cover heuristic removed (S11): fully subsumed by M1 color-block above

        # M1 heuristic (extended): low information-density badges.
        # CC license badges, journal logos, and similar decorative
        # images have very low bytes-per-pixel (< 0.15) AND few
        # unique colors (< 50).  Real content (photos, charts,
        # diagrams) has either high bytes-per-pixel or many colors.
        # This catches CC-BY badges (~0.07 B/px, ~35 colors) that
        # pass other heuristics due to moderate dimensions.
        # Uses file_size from the size check at the top of this function.
        #
        # F3 Safeguard 2: skip M1-extended for large files (> 50KB).
        # soffice chart renders are large PNGs (100KB+) with white
        # backgrounds and thin colored curves. Their bpp is low because
        # PNG compression is efficient on white space, but they are real
        # substantive images. The 50KB gate is well above real blank
        # artifacts (< 10KB) and well below typical chart renders (50KB+).
        try:
            _orig_w, _orig_h = img.size  # original dimensions
            _pixel_count = _orig_w * _orig_h
            if (_pixel_count > 0 and unique_colors >= 0
                    and file_size > 0
                    and file_size <= 50000):  # F3 Safeguard 2
                _bpp = file_size / _pixel_count
                if _bpp < 0.15 and unique_colors < 50 and gray_std < 40:
                    return True
        except Exception:
            pass

    except Exception:
        pass  # If PIL/numpy unavailable, skip this check

    return False


def _is_journal_branding(img_detail: dict,
                         file_size_bytes: int = 0,
                         page_num: int = 0,
                         page_image_count: int = 0) -> bool:
    """FIX-3 (S36 enhanced): Detect journal branding elements.

    Identifies common decorative elements in journal PDFs:
      - Check 1: Very small images (width < 100px AND height < 100px)
      - Check 2: Extreme aspect ratios (> 10:1 or < 1:10) -- banners/bars
      - Check 3: Tiny file size (< 5KB) with relaxed dimension gate
      - Check 4: Page-1 moderate images on multi-image pages (NEW S36)
      - Check 5: Page-1 high-image-count website extraction (NEW S36)
      - Check 6: Page-1 small images regardless of count (NEW S36)

    Args:
        img_detail: Dict with width, height keys.
        file_size_bytes: File size in bytes (0 = unknown/skip check).
        page_num: Page number (1-indexed). 0 = unknown. (S36)
        page_image_count: Total images on this page. 0 = unknown. (S36)

    Returns True if the image is likely journal branding / decorative.
    """
    w = img_detail.get("width", 0)
    h = img_detail.get("height", 0)

    # Skip if dimensions are unknown
    if w <= 0 or h <= 0:
        return False

    # S36: Check 1: Very small images (icons, badges, bullets) [unchanged]
    # BOTH dimensions must be small to avoid false positives on
    # narrow-but-tall images (gel lanes ~80x400px) or wide-but-short
    # images (spectral traces ~500x60px).
    if w < 100 and h < 100:
        return True

    # S36: Check 2: Extreme aspect ratio (banners, colored bars, dividers) [unchanged]
    # GATED on file_size < 5KB: legitimate wide/tall process diagrams
    # can have extreme aspect ratios but are typically > 5KB file size.
    # When file_size is unknown (0), still apply the filter (conservative).
    aspect = w / h if h > 0 else 0
    if aspect > 10.0 or aspect < 0.1:
        if file_size_bytes == 0 or file_size_bytes < 5000:
            return True
        # Large file with extreme aspect: likely a real diagram, do not filter

    # S36: Check 3 (RELAXED): Tiny file size (< 5KB)
    # OLD: w < 200 AND h < 200 (missed SAGE logo 669x219 at 4KB)
    # S36-FIX: Page 1 uses relaxed OR gate (catches wide logos like SAGE
    # 669x219 at 4KB). Other pages use original AND gate to avoid killing
    # legitimate small diagrams (e.g. 500x300 sparkline at 4.5KB).
    if 0 < file_size_bytes < 5000:
        if page_num == 1:
            # S36-FIX: CRITICAL-01 - Relaxed gate for page 1 only
            if w < 400 or h < 400:
                return True
        else:
            # S36-FIX: CRITICAL-01 - Original AND gate for non-page-1
            if w < 200 and h < 200:
                return True

    # S36: Check 4 (NEW): Page-1 small-to-medium images on multi-image pages
    # On page 1 with >3 images, images under ~50KB with moderate dimensions
    # are very likely journal/publisher branding. Real paper figures are
    # typically >50KB and appear on later pages.
    if page_num == 1 and page_image_count > 3:
        if 0 < file_size_bytes < 50000:
            if w < 500 and h < 500:
                return True

    # S36: Check 5 (NEW): Page-1 high-image-count = journal website extraction
    # When page 1 has >8 images, almost all are decorative website elements.
    # Only very large images (>100KB) should escape this filter.
    # S36-FIX: MAJOR-01 - Added > 0 guard. When file_size_bytes is 0
    # (unknown), 0 < 100000 would be TRUE, falsely classifying all
    # unknown-size images as branding.
    if page_num == 1 and page_image_count > 8:
        if 0 < file_size_bytes < 100000:
            return True

    # S36: Check 6 (NEW): Page-1 small images regardless of image count
    # Journal covers and publisher logos on page 1 are typically <300x300
    # and <50KB. A real paper figure at this size would be unusual.
    if page_num == 1:
        if w < 300 and h < 300 and 0 < file_size_bytes < 50000:
            return True

    return False


def _classify_single_image(
    img,         # type: dict
    page_data,   # type: dict
    xref_counts, # type: dict
    total_pages, # type: int
):
    # type: (...) -> None
    """Classify a single image as substantive or decorative.

    Sets img["is_substantive"] (bool) and img["classification_reason"] (str)
    in-place. Implements an ordered heuristic chain:
      [1] Blank image check (file-level)
      [2] H1: Tiny dimensions
      [3] H9/FIX-3: Journal branding
      [4] H4: Repeated xref (watermark)
      [5] H-NEW: Low word count for vector renders
      [6] H5: Large image with text context
      [7] H7: Figure keywords in context
      [8] Conservative fallback (SUB)
    """
    import os as _os_classify

    context = page_data.get("context", "").lower()
    w = img.get("width", 0)
    h = img.get("height", 0)
    xref = img.get("xref")
    file_path = img.get("file_path", "")

    # [1] FILE-LEVEL BLANK CHECK
    if file_path:
        if _is_blank_image(file_path):
            img["is_substantive"] = False
            img["classification_reason"] = "blank_image"
            return

    # [1b] NARROW FRAGMENT CHECK (m3 fix)
    # Images with height < 50px or width < 50px are narrow table fragments
    # or decorative strips (e.g. 282x36px bottom row of a spreadsheet).
    # Checked separately from H1 because H1 requires BOTH dims < 50.
    if (h > 0 and h < 50 and w >= 50) or (w > 0 and w < 50 and h >= 50):
        img["is_substantive"] = False
        img["classification_reason"] = (
            f"narrow_fragment (h={h}px)" if h < 50
            else f"narrow_fragment (w={w}px)"
        )
        return

    # [2] H1: TINY DIMENSIONS
    if w < 50 and h < 50 and w > 0 and h > 0:
        img["is_substantive"] = False
        img["classification_reason"] = "tiny_dimensions"
        return

    # [3] H9/FIX-3: JOURNAL BRANDING (S36 enhanced with page context)
    _file_size = 0
    if file_path:
        try:
            _file_size = _os_classify.path.getsize(file_path)
        except OSError:
            pass
    # S36: Extract page-level info for branding check
    _page_num = page_data.get("page", 0)
    _page_image_count = page_data.get("image_count", 0)
    if _is_journal_branding(img, file_size_bytes=_file_size,
                            page_num=_page_num,
                            page_image_count=_page_image_count):
        img["is_substantive"] = False
        img["classification_reason"] = "journal_branding"
        return

    # [4] H4: REPEATED XREF (watermark)
    if xref is not None and total_pages > 0:
        if xref_counts.get(xref, 0) > total_pages * 0.5:
            img["is_substantive"] = False
            img["classification_reason"] = "repeated_watermark"
            return

    # [5] H-NEW: LOW WORD COUNT FOR VECTOR RENDERS
    # Section-title slides rendered as vector have < 20 words of text.
    if img.get("source") == "vector-render":
        _page_text = page_data.get("full_text",
                                    page_data.get("context", ""))
        _word_count = len(_page_text.split())
        if _word_count < 20:
            img["is_substantive"] = False
            img["classification_reason"] = "low_word_count_render"
            return

    # [6] H5: LARGE IMAGE WITH TEXT CONTEXT
    if w > 200 and h > 200 and len(context) > 20:
        img["is_substantive"] = True
        img["classification_reason"] = "large_with_context"
        return

    # [7] H7: FIGURE KEYWORDS
    figure_keywords = ["figure", "table", "diagram", "model",
                       "chart", "graph", "plot", "curve",
                       "illustration", "schematic"]
    if any(kw in context for kw in figure_keywords):
        img["is_substantive"] = True
        img["classification_reason"] = "figure_keyword"
        return

    # [8] CONSERVATIVE FALLBACK
    img["is_substantive"] = True
    img["classification_reason"] = "conservative_default"


def _classify_page_images(page_data: dict,
                          all_pages: list,
                          page_count: int,
                          xref_counts=None) -> bool:
    """Classify whether a page's images are substantive or decorative.

    Returns True if substantive, False if decorative.
    Conservative: when uncertain, classify as substantive.

    Args:
        page_data: Dict for the current page.
        all_pages: List of all page dicts (used to build xref_counts
                   if not provided).
        page_count: Total number of pages in the document.
        xref_counts: Optional pre-built dict mapping xref -> count
                     across all pages.  When None, built internally
                     (O(I) per call; pass pre-built to avoid O(P*I)).

    Filtering heuristics (from requirements):
      - Image dimensions < 50x50 px -> Decorative
      - Page 1 AND context suggests title/cover -> Decorative
      - Last page AND context contains "thank" or "question" -> Decorative
      - Same xref on >50% of pages -> Decorative (watermarks/headers)
      - Image > 200x200 px AND page has >20 chars text -> Substantive
      - Context contains figure keywords -> Substantive
    """
    page_num = page_data["page"]
    images = page_data.get("image_details", [])
    context = page_data.get("context", "").lower()
    image_count = page_data["image_count"]

    # Heuristic 8 (checked FIRST): vector content detection for pure-vector pages.
    # Pure-vector pages (e.g. Kaplan-Meier curves, SmartArt) have image_count == 0
    # because pymupdf does not report vector paths as raster images.  If we let the
    # image_count == 0 short-circuit run first, these pages are always classified as
    # decorative and never reach this check.  Moved before the short-circuit so that
    # any page with significant vector content is immediately classified as substantive.
    # BUG-3 fix: uses combined drawing_count + bounding-box area heuristic.
    if _has_significant_vector_content(page_data):
        drawing_count = page_data.get("drawing_count", 0)
        _warn(f"Page {page_num}: vector content detected "
              f"({drawing_count} drawings, "
              f"max_area={page_data.get('max_drawing_area_pct', 0):.1f}%) "
              f"— classifying as substantive")
        # Mark all images on this page as SUB for per-image tracking
        for img in images:
            img["is_substantive"] = True
            img["classification_reason"] = "vector_page"
        return True  # Substantive vector content detected

    if image_count == 0:
        return False  # No raster images and no vector content — truly empty

    # Build xref frequency map across all pages for watermark detection
    # When xref_counts is pre-built by the caller, skip the O(I) rebuild.
    if xref_counts is None:
        xref_counts = {}
        for p in all_pages:
            for img in p.get("image_details", []):
                xref = img.get("xref")
                if xref is not None:
                    xref_counts[xref] = xref_counts.get(xref, 0) + 1

    # Per-image classification via _classify_single_image()
    for img in images:
        _classify_single_image(img, page_data, xref_counts, page_count)

    has_substantive = any(
        img.get("is_substantive", False) for img in images)
    all_decorative = all(
        not img.get("is_substantive", False) for img in images)

    # Heuristic 2: page 1 with title/cover context
    if page_num == 1 and not has_substantive:
        title_keywords = ["title", "cover", "university", "logo",
                          "department", "faculty"]
        if any(kw in context for kw in title_keywords):
            return False

    # Heuristic 3: last page with "thank" or "question"
    if page_num == page_count and not has_substantive:
        if "thank" in context or "question" in context:
            return False

    # Conservative default: if we have images that weren't all filtered out
    # and no strong decorative signal, classify as substantive
    return has_substantive or not all_decorative


def scan_pdf_images(pdf_path: str) -> Optional[list]:
    """Scan a PDF and return per-page image data.

    R19: Uses pymupdf (fitz) to detect images per page.
    Returns list of dicts with page, image_count, context, image_details.
    Returns [] if fitz is not installed (library availability issue).
    Returns None if the document cannot be opened (encrypted/corrupt),
    triggering the error manifest path in generate_image_index().
    """
    if not _HAS_FITZ:
        _warn("fitz (PyMuPDF) not installed. Cannot scan PDF for images.")
        return []

    try:
        doc = fitz.open(str(pdf_path))
    except Exception as e:
        _warn(f"Could not open PDF for image scanning: {e}")
        return None

    pages = []
    for page_num in range(len(doc)):
        try:
            page = doc[page_num]
            images = page.get_images(full=True)

            # Heuristic 8: count vector drawing paths for SmartArt detection
            # BUG-3 fix: also compute max drawing bounding-box area as a
            # percentage of the page area.  This distinguishes real figures
            # (few large shapes) from table styling (many small shapes).
            max_drawing_area_pct = 0.0
            try:
                drawings = page.get_drawings()
                drawing_count = len(drawings)
                if drawing_count >= VECTOR_DRAWING_THRESHOLD_LOW:
                    # Compute area for all pages with enough drawings
                    page_rect = page.rect
                    page_area = page_rect.width * page_rect.height
                    if page_area > 0:
                        for d in drawings:
                            r = d.get("rect")
                            if r is not None:
                                d_area = abs(r.width * r.height)
                                d_pct = (d_area / page_area) * 100.0
                                if d_pct > max_drawing_area_pct:
                                    max_drawing_area_pct = d_pct
            except Exception:
                drawing_count = 0

            text = page.get_text().strip()
            # Store first 500 chars for word count heuristic (H-NEW)
            full_text = text[:500]

            # Extract first meaningful line (>10 chars, not just numbers)
            context = ""
            for line in text.split("\n"):
                line = line.strip()
                if len(line) >= 10 and not line.isdigit():
                    context = line[:150]
                    if len(line) > 150:
                        context += "..."
                    break

            pages.append({
                "page": page_num + 1,  # 1-indexed
                "image_count": len(images),
                "drawing_count": drawing_count,
                "max_drawing_area_pct": max_drawing_area_pct,
                "context": context,
                "full_text": full_text,
                "image_details": [
                    {
                        "xref": img[0],
                        "width": img[2],
                        "height": img[3],
                        "colorspace": str(img[5]) if len(img) > 5 else "",
                    }
                    for img in images
                ]
            })
        except Exception as e:
            _warn(f"Could not scan page {page_num + 1}: {e}")
            pages.append({
                "page": page_num + 1,
                "image_count": 0,
                "drawing_count": 0,
                "max_drawing_area_pct": 0.0,
                "context": "",
                "full_text": "",
                "image_details": [],
            })

    doc.close()
    return pages


def _build_xref_filepath_map(
    images_dir,  # type: Optional[Path]
    pages,       # type: list
):
    # type: (...) -> dict
    """Build a mapping from (page_num, image_index) to file path on disk.

    The extracted image naming convention from convert-paper.py is:
        page{N}-img{M}.{ext}   (e.g., page5-img1.png, page5-img2.jpeg)
    Also handles: fig{N}-page{M}.{ext} and page{N}-vector-render.png

    Populates image_details[N]["file_path"] in-place.
    Returns the mapping dict for reference.
    """
    if images_dir is None or not images_dir.exists():
        return {}

    # Build inventory of files on disk by page number
    # Pattern: page{N}-img{M}.ext  OR  page{N}-vector-render.ext
    file_map = {}   # (page_num, img_index) -> file_path
    page_files = {}  # page_num -> [file_paths] (sorted)

    _img_exts = {".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".tif", ".gif"}
    for f in sorted(images_dir.iterdir()):
        if not f.is_file() or f.suffix.lower() not in _img_exts:
            continue
        # Try to extract page number from filename
        pg_match = re.search(
            r'page[\-_]?(\d+)', f.stem, re.IGNORECASE)
        if not pg_match:
            pg_match = re.search(
                r'(?:^|[\-_])p(\d+)(?:[\-_]|$)',
                f.stem, re.IGNORECASE)
        if pg_match:
            pg_num = int(pg_match.group(1))
            page_files.setdefault(pg_num, []).append(f)

    # Match files to image_details by page and order
    for page_data in pages:
        pg_num = page_data["page"]
        files_for_page = page_files.get(pg_num, [])
        images = page_data.get("image_details", [])

        # Vector renders: if page has vector_rendered flag, the single
        # file is the vector render (find it by name pattern)
        if page_data.get("vector_rendered"):
            for _vi, img in enumerate(images):
                if img.get("source") == "vector-render":
                    # Find the render file
                    for f in files_for_page:
                        if "vector-render" in f.stem or "render" in f.stem:
                            img["file_path"] = str(f)
                            file_map[(pg_num, _vi)] = str(f)
                            break
            continue

        # Raster images: match by position (same order as fitz.get_images)
        # Filter out vector renders from the file list
        raster_files = [
            f for f in files_for_page
            if "vector-render" not in f.stem and "render" not in f.stem
        ]

        for i, img in enumerate(images):
            if i < len(raster_files):
                img["file_path"] = str(raster_files[i])
                file_map[(pg_num, i)] = str(raster_files[i])

    # Drop ghost entries: xref records with no matching file on disk.
    # Entries with empty/missing file_path were not matched to any
    # raster file and are ghost xref records.
    for page_data in pages:
        valid_details = []
        for img in page_data.get("image_details", []):
            fp = img.get("file_path", "")
            if fp and Path(fp).exists():
                valid_details.append(img)
        page_data["image_details"] = valid_details
        page_data["image_count"] = len(valid_details)

    return file_map


# ═══════════════════════════════════════════════════════════════════════════
# m3: IMAGE INDEX OVERRIDES (image-index-overrides.json)
# ═══════════════════════════════════════════════════════════════════════════
# Allows manual correction of automatic SUBSTANTIVE/DECORATIVE
# classifications.  An override file can be placed alongside the source
# document or in the --target-dir.  If no file is found, the pipeline
# behaves identically to before (silent no-op).
#
# Override file format:
# {
#   "overrides": [
#     {
#       "file_pattern": "*.pptx",
#       "pages": {
#         "5": {"classification": "SUBSTANTIVE", "reason": "Key diagram"},
#         "12": {"classification": "DECORATIVE", "reason": "Background"}
#       }
#     }
#   ]
# }
# ═══════════════════════════════════════════════════════════════════════════

_OVERRIDE_FILENAME = "image-index-overrides.json"


def _load_image_index_overrides(
    source_path: Path,
    target_dir: Optional[Path] = None,
) -> Optional[list]:
    """m3: Load image-index-overrides.json if it exists.

    Search order:
      1. Same directory as the source file.
      2. The --target-dir (if specified and different from source dir).

    Returns the "overrides" list from the JSON, or None if no file
    found or the file is invalid.  Invalid JSON logs a warning and
    returns None (pipeline continues without overrides).
    """
    candidates = [source_path.parent / _OVERRIDE_FILENAME]
    if target_dir is not None:
        resolved_target = target_dir.resolve()
        resolved_source_dir = source_path.parent.resolve()
        if resolved_target != resolved_source_dir:
            candidates.append(resolved_target / _OVERRIDE_FILENAME)

    for candidate in candidates:
        if candidate.exists():
            try:
                data = json.loads(candidate.read_text(encoding="utf-8"))
                overrides = data.get("overrides")
                if not isinstance(overrides, list):
                    _warn(f"m3: Override file {candidate} has no valid "
                          "'overrides' list. Ignoring.")
                    return None
                print(f"  m3: Loaded overrides from {candidate}")
                print(f"      {len(overrides)} file pattern(s) defined")
                return overrides
            except json.JSONDecodeError as e:
                _warn(f"m3: Invalid JSON in {candidate}: {e}. "
                      "Continuing without overrides.")
                return None
            except Exception as e:
                _warn(f"m3: Could not read {candidate}: {e}. "
                      "Continuing without overrides.")
                return None

    # No override file found — silent, expected behavior
    return None


def _apply_overrides_to_pages(
    pages: list,
    overrides: list,
    source_filename: str,
) -> int:
    """m3: Apply overrides to page classification data (in-place).

    For each override entry whose file_pattern matches source_filename
    (using fnmatch), applies per-page classification overrides.

    Args:
        pages: List of page dicts with 'page' (1-indexed) and
               'is_substantive' (bool) keys.  Modified in-place.
        overrides: The "overrides" list from the JSON file.
        source_filename: Filename (not full path) of the source document.

    Returns:
        Number of pages that were overridden.
    """
    from fnmatch import fnmatch

    override_count = 0

    for entry in overrides:
        file_pattern = entry.get("file_pattern", "")
        if not fnmatch(source_filename, file_pattern):
            continue

        page_overrides = entry.get("pages", {})
        if not isinstance(page_overrides, dict):
            _warn(f"m3: 'pages' for pattern '{file_pattern}' is not "
                  "a dict. Skipping this entry.")
            continue

        for page_str, page_config in page_overrides.items():
            try:
                target_page = int(page_str)
            except (ValueError, TypeError):
                _warn(f"m3: Invalid page number '{page_str}' in "
                      f"pattern '{file_pattern}'. Skipping.")
                continue

            if not isinstance(page_config, dict):
                _warn(f"m3: Config for page {target_page} in pattern "
                      f"'{file_pattern}' is not a dict. Skipping.")
                continue

            classification = page_config.get("classification", "").upper()
            reason = page_config.get("reason", "no reason given")

            if classification not in ("SUBSTANTIVE", "DECORATIVE"):
                _warn(f"m3: Invalid classification '{classification}' "
                      f"for page {target_page} in pattern "
                      f"'{file_pattern}'. Must be SUBSTANTIVE or "
                      f"DECORATIVE. Skipping.")
                continue

            new_is_substantive = (classification == "SUBSTANTIVE")

            # Find the matching page in the pages list
            for page_data in pages:
                if page_data["page"] == target_page:
                    old_val = page_data.get("is_substantive", False)
                    old_label = ("SUBSTANTIVE" if old_val
                                 else "DECORATIVE")
                    new_label = classification

                    if old_val != new_is_substantive:
                        page_data["is_substantive"] = new_is_substantive
                        override_count += 1
                        print(f"      Page {target_page}: "
                              f"{old_label} -> {new_label} "
                              f"(reason: {reason})")
                    else:
                        print(f"      Page {target_page}: already "
                              f"{new_label}, no change "
                              f"(reason: {reason})")
                    break
            else:
                _warn(f"m3: Page {target_page} not found in document "
                      f"(pattern '{file_pattern}'). Skipping.")

    return override_count


def _apply_overrides_to_image_index_file(
    index_path: Path,
    overrides: list,
    source_filename: str,
) -> int:
    """m3: Apply overrides to an existing image index .md file on disk.

    Used for PPTX/DOCX formats where convert-office.py generates the
    image index.  Reads the markdown table, applies classification
    overrides, rewrites both the Page-by-Page table and the
    Substantive Images Only table, and updates summary counts.

    Args:
        index_path: Path to the *-image-index.md file.
        overrides: The "overrides" list from the JSON file.
        source_filename: Filename of the source document.

    Returns:
        Number of pages that were overridden.
    """
    from fnmatch import fnmatch

    # Collect all page overrides that match this file
    page_overrides_map: dict = {}  # page_num -> (classification, reason)
    for entry in overrides:
        file_pattern = entry.get("file_pattern", "")
        if not fnmatch(source_filename, file_pattern):
            continue
        page_overrides = entry.get("pages", {})
        if not isinstance(page_overrides, dict):
            continue
        for page_str, page_config in page_overrides.items():
            try:
                page_num = int(page_str)
            except (ValueError, TypeError):
                continue
            if not isinstance(page_config, dict):
                continue
            classification = page_config.get(
                "classification", "").upper()
            if classification not in ("SUBSTANTIVE", "DECORATIVE"):
                continue
            reason = page_config.get("reason", "no reason given")
            page_overrides_map[page_num] = (classification, reason)

    if not page_overrides_map:
        return 0

    # Read and modify the image index file
    try:
        content = index_path.read_text(encoding="utf-8")
    except Exception as e:
        _warn(f"m3: Could not read image index {index_path}: {e}")
        return 0

    lines = content.split("\n")
    override_count = 0
    new_lines = []
    # Track counts for updating the filtering summary
    substantive_delta = 0
    # Collect all substantive page data for rebuilding the
    # "Substantive Images Only" table after overrides.
    # Format: list of (page_num, image_count, context) tuples.
    all_page_data: list = []  # populated from Page-by-Page table
    # Track which section we are in to handle the substantive table
    in_substantive_section = False
    substantive_header_seen = False
    substantive_separator_seen = False
    skip_old_substantive_rows = False
    subst_col_count = 3  # Default: PDF uses 3 cols; office uses 4

    for line in lines:
        # Detect the "## Substantive Images Only" section
        if line.strip() == "## Substantive Images Only":
            in_substantive_section = True
            new_lines.append(line)
            continue

        # Detect the "## Filtering Summary" section (ends substantive)
        if line.strip() == "## Filtering Summary":
            in_substantive_section = False
            skip_old_substantive_rows = False
            new_lines.append(line)
            continue

        # Inside the substantive section: skip old data rows
        # (they will be rebuilt below).  Keep headers and separators.
        if in_substantive_section:
            if line.startswith("|"):
                cols = [c.strip() for c in line.split("|")]
                actual_cols = cols[1:-1] if len(cols) > 2 else []
                # Header row: contains "Page" text
                if not substantive_header_seen and any(
                        "Page" in c for c in actual_cols):
                    substantive_header_seen = True
                    # Detect column count from header for rebuild
                    subst_col_count = len(actual_cols)
                    new_lines.append(line)
                    continue
                # Separator row: all dashes
                if (not substantive_separator_seen
                        and substantive_header_seen
                        and all(c.replace("-", "").strip() == ""
                                for c in actual_cols)):
                    substantive_separator_seen = True
                    skip_old_substantive_rows = True
                    new_lines.append(line)
                    # Rebuild substantive rows matching the detected
                    # column count.  PDF indexes use 3 cols
                    # (Page | Images | Context); office indexes use 4
                    # (Page | Images | Context | Notes).
                    for pg_entry in all_page_data:
                        if len(pg_entry) == 4:
                            pg, img_ct, ctx, notes = pg_entry
                        else:
                            pg, img_ct, ctx = pg_entry
                            notes = ""
                        if subst_col_count >= 4:
                            new_lines.append(
                                f"| {pg} | {img_ct} | {ctx} "
                                f"| {notes} |")
                        else:
                            new_lines.append(
                                f"| {pg} | {img_ct} | {ctx} |")
                    if not all_page_data:
                        if subst_col_count >= 4:
                            new_lines.append(
                                "| — | — | No substantive images "
                                "found |  |")
                        else:
                            new_lines.append(
                                "| — | — | No substantive images "
                                "found |")
                    continue
                # Data row in old substantive table: skip it
                if skip_old_substantive_rows:
                    continue
            else:
                # Non-table line in substantive section (e.g. blank,
                # "---"): stop skipping and pass through
                skip_old_substantive_rows = False
                new_lines.append(line)
                continue

        # Process Page-by-Page table rows (4 or 5 column table).
        # The table format may have 4 cols (Page|Images|Substantive|Context)
        # or 5 cols (Page|Images|Substantive|Context|Notes).  Both are valid.
        if line.startswith("|") and "|" in line[1:]:
            cols = [c.strip() for c in line.split("|")]
            actual_cols = cols[1:-1] if len(cols) > 2 else []

            if len(actual_cols) >= 4:
                try:
                    page_num = int(actual_cols[0])
                except (ValueError, TypeError):
                    new_lines.append(line)
                    continue

                if page_num in page_overrides_map:
                    classification, reason = page_overrides_map[
                        page_num]
                    old_val = actual_cols[2]  # "Yes" / "No" / "No (duplicate)"
                    new_val = ("Yes" if classification == "SUBSTANTIVE"
                               else "No")
                    # Normalise old_val for comparison: any "No*" counts as
                    # not-substantive, "Yes" counts as substantive.
                    old_is_subst = old_val.strip().startswith("Yes")
                    new_is_subst = (new_val == "Yes")

                    if old_is_subst != new_is_subst:
                        actual_cols[2] = new_val
                        override_count += 1
                        if new_is_subst:
                            substantive_delta += 1
                        else:
                            substantive_delta -= 1
                        old_label = ("SUBSTANTIVE" if old_is_subst
                                     else "DECORATIVE")
                        print(f"      Page {page_num}: "
                              f"{old_label}"
                              f" -> {classification} "
                              f"(reason: {reason})")
                        line = "| " + " | ".join(actual_cols) + " |"
                    else:
                        print(f"      Page {page_num}: already "
                              f"{classification}, no change "
                              f"(reason: {reason})")

                # Collect substantive pages for rebuilding the
                # substantive-only table.  Store (page, images, context,
                # notes) for 5-col rows or (page, images, context) for
                # 4-col rows so the rebuild can emit correct column count.
                is_subst = actual_cols[2].strip()
                if is_subst == "Yes":
                    _notes = actual_cols[4] if len(actual_cols) >= 5 else ""
                    all_page_data.append(
                        (actual_cols[0], actual_cols[1],
                         actual_cols[3], _notes))

            new_lines.append(line)
        else:
            new_lines.append(line)

    # Update the "Estimated substantive images" count if changed
    if substantive_delta != 0:
        for i, nline in enumerate(new_lines):
            if nline.startswith("Estimated substantive images:"):
                try:
                    parts = nline.split(":")
                    count_part = parts[1].strip()
                    old_count = int(count_part.split()[0])
                    new_count = max(0, old_count + substantive_delta)
                    new_lines[i] = (
                        f"Estimated substantive images: {new_count} "
                        f"(after filtering + {override_count} "
                        f"override(s))"
                    )
                except (ValueError, IndexError):
                    pass
                break

    # Update filtering summary: add override note
    if override_count > 0:
        for i, nline in enumerate(new_lines):
            if nline.startswith("- Filtering criteria applied:"):
                # Insert override note after the filtering criteria line
                new_lines.insert(
                    i + 1,
                    f"- Manual overrides applied: {override_count} "
                    f"page(s) (via image-index-overrides.json)")
                break

    # Also update decorative/substantive counts in filtering summary
    if substantive_delta != 0:
        for i, nline in enumerate(new_lines):
            if nline.startswith("- Pages classified as decorative:"):
                try:
                    old_dec = int(nline.split(":")[1].strip())
                    new_lines[i] = (
                        f"- Pages classified as decorative: "
                        f"{max(0, old_dec - substantive_delta)}")
                except (ValueError, IndexError):
                    pass
            elif nline.startswith(
                    "- Pages classified as substantive:"):
                try:
                    old_sub = int(nline.split(":")[1].strip())
                    new_lines[i] = (
                        f"- Pages classified as substantive: "
                        f"{max(0, old_sub + substantive_delta)}")
                except (ValueError, IndexError):
                    pass

    # Write updated content back
    if override_count > 0:
        try:
            index_path.write_text("\n".join(new_lines), encoding="utf-8")
            print(f"      Rewrote {index_path.name} with "
                  f"{override_count} override(s)")
        except Exception as e:
            _warn(f"m3: Could not rewrite {index_path}: {e}")
            return 0

    return override_count


def _update_manifest_with_vector_renders(
        output_md: Path,
        vector_rendered: list,
        images_dir: Path,
) -> None:
    """Add vector-rendered images to the image manifest JSON.

    Called by generate_image_index() after rendering pure-vector pages.
    Appends entries to the existing manifest or creates a minimal one.
    """
    # Find existing manifest
    manifest_candidates = [
        images_dir / "image-manifest.json",
        output_md.parent / f"{output_md.stem}_manifest.json",
    ]
    manifest_path = next(
        (p for p in manifest_candidates if p.exists()), None)

    if manifest_path is not None:
        try:
            data = json.loads(
                manifest_path.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, ValueError):
            data = {"images": [], "image_count": 0,
                    "total_images": 0,
                    "images_dir": str(images_dir),
                    "md_file": str(output_md),
                    "source": str(output_md.parent / output_md.stem)}
    else:
        manifest_path = images_dir / "image-manifest.json"
        data = {"images": [], "image_count": 0,
                "total_images": 0,
                "images_dir": str(images_dir),
                "md_file": str(output_md),
                "source": str(output_md.parent / output_md.stem)}

    existing_filenames = {
        img.get("filename") for img in data.get("images", [])}

    added = 0
    base_fig = max(
        (fn for fn in
         (img.get("figure_num", 0) for img in data.get("images", []))
         if isinstance(fn, (int, float))),
        default=0,
    )
    for idx, vr in enumerate(vector_rendered, start=1):
        if vr["filename"] in existing_filenames:
            continue
        entry = {
            "index": len(data.get("images", [])),
            "figure_num": base_fig + idx,
            "filename": vr["filename"],
            "file_path": vr["file_path"],
            "width": vr["width"],
            "height": vr["height"],
            "page": vr["page"],
            "source": "vector-render",
            "drawing_count": vr["drawing_count"],
            "decorative": False,
            "blank": False,
            "is_duplicate": False,
            "analysis_status": "pending",
            "description": (f"Vector-rendered page {vr['page']} "
                            f"({vr['drawing_count']} drawings)"),
            "type_guess": "vector_render",
            "section_context": "",
            "detected_caption": None,
        }
        data.setdefault("images", []).append(entry)
        added += 1

    if added > 0:
        data["image_count"] = len(data.get("images", []))
        data["total_images"] = data["image_count"]
        try:
            manifest_path.write_text(
                json.dumps(data, indent=2, ensure_ascii=False) + "\n",
                encoding="utf-8")
            print(f"  FIX-1: Added {added} vector render(s) to "
                  f"{manifest_path.name}")
        except Exception as e:
            _warn(f"Could not update manifest with vector "
                  f"renders: {e}")


def _generate_image_index_from_mineru_manifest(
    source_path: Path,
    output_md: Path,
    index_path: Path,
    images_dir: Optional[Path],
    target_dir: Optional[Path],
    extractor_label: str = "MinerU",
) -> Optional[dict]:
    """Generate image index from image-manifest.json.

    Called by generate_image_index() when extractor == "mineru" OR
    when extractor == "pymupdf4llm" and image-manifest.json exists
    (RC-B fix: manifest is the authoritative source for all extractors,
    including split-panel images created by _split_panels in
    convert-paper.py that are missed by the fitz xref scan).

    Reads the manifest and produces a companion *-image-index.md file
    that is consistent with the actual images on disk.

    Args:
        source_path: Path to the original PDF.
        output_md: Path to the converted .md file.
        index_path: Path for the output image index .md file.
        images_dir: Path to the images directory containing
                    image-manifest.json.
        target_dir: Optional --target-dir for m3 overrides.
        extractor_label: Display label for the extractor in printed
                         output and the written index file.

    Returns:
        Dict with image index metadata for R21 registry, or None
        on failure.
    """
    print(f"  Extractor: {extractor_label} "
          f"(reading manifest, not scanning PDF)")

    # ── 1. Locate and read the MinerU manifest ──
    manifest_path = None
    if images_dir is not None:
        candidate = images_dir / "image-manifest.json"
        if candidate.exists():
            manifest_path = candidate
    if manifest_path is None:
        # Fallback: look alongside the output .md
        for _d in [output_md.parent / "images",
                   output_md.parent / f"{output_md.stem}_images"]:
            _c = _d / "image-manifest.json"
            if _c.exists():
                manifest_path = _c
                break
    if manifest_path is None:
        _warn("MinerU image index: image-manifest.json not found. "
              "Cannot generate image index.")
        _write_error_image_index(
            index_path, source_path,
            "MinerU manifest not found — cannot generate image index")
        return {
            "image_index_path": str(index_path),
            "image_index_generated_at": (
                datetime.now(timezone.utc).isoformat()),
            "total_pages": 0,
            "pages_with_images": 0,
            "total_images_detected": 0,
            "substantive_images": 0,
            "has_testable_images": False,
        }

    try:
        manifest_data = json.loads(
            manifest_path.read_text(encoding="utf-8"))
    except Exception as e:
        _warn(f"MinerU image index: could not read manifest: {e}")
        return None

    manifest_images = manifest_data.get("images", [])
    total_images = len(manifest_images)
    print(f"  Manifest: {total_images} images from "
          f"{manifest_path.name}")

    # ── 2. Build per-page data from manifest entries ──
    # Group images by page number for the page-by-page table.
    page_map = {}  # page_num -> list of image entries
    for img in manifest_images:
        pg = img.get("page") or 0
        page_map.setdefault(pg, []).append(img)

    # Get total page count from YAML frontmatter if available
    total_pages = 0
    try:
        md_text = output_md.read_text(encoding="utf-8")
        import re as _re_mu
        _pages_match = _re_mu.search(
            r'^pages:\s*(\d+)', md_text, _re_mu.MULTILINE)
        if _pages_match:
            total_pages = int(_pages_match.group(1))
    except Exception:
        pass
    if total_pages == 0:
        # Fallback: highest page number in manifest
        total_pages = max(
            (img.get("page") or 0 for img in manifest_images),
            default=0)

    pages_with_images = len([pg for pg in page_map if pg > 0])

    # ── 3. Classify images ──
    # MinerU manifest already flags near_black_detected.
    # For classification, treat near-black (not re-rendered) as
    # decorative.  All other images are conservatively substantive
    # unless they are very small (likely icons/logos).
    import os as _os_mineru
    substantive_count = 0
    for img in manifest_images:
        is_sub = True
        _src = img.get("mineru_source", "images")
        _extr_src = img.get("extraction_source", "mineru")
        reason = (f"table screenshot (MinerU tables/)"
                  if _src == "tables"
                  else "fitz fallback extraction"
                  if _extr_src == "fitz_fallback"
                  else "default (MinerU extraction)")
        # Near-black that failed re-render = decorative
        if img.get("near_black_detected") and not img.get("rerendered"):
            is_sub = False
            reason = "near-black (render failed)"
        # Narrow fragment: one dimension < 50px, other >= 50px (m3 fix)
        # e.g. 282x36px = bottom row of a spreadsheet, not a real image.
        elif (img.get("width", 0) > 0 and img.get("height", 0) > 0
              and ((img["height"] < 50 and img["width"] >= 50)
                   or (img["width"] < 50 and img["height"] >= 50))):
            is_sub = False
            w_val = img["width"]
            h_val = img["height"]
            reason = (f"narrow_fragment (h={h_val}px)" if h_val < 50
                      else f"narrow_fragment (w={w_val}px)")
        # Very small images (< 50x50 or < 2KB) = likely decorative
        elif (img.get("width", 0) > 0 and img.get("height", 0) > 0
              and img["width"] < 50 and img["height"] < 50):
            is_sub = False
            reason = "tiny image (< 50x50)"
        elif img.get("file_path"):
            try:
                fsize = _os_mineru.path.getsize(img["file_path"])
                if fsize < 2048:
                    is_sub = False
                    reason = "tiny file (< 2KB)"
            except OSError:
                pass
        img["is_substantive"] = is_sub
        img["classification_reason"] = reason
        if is_sub:
            substantive_count += 1

    # ── 4. Apply m3 overrides if present ──
    _m3_overrides = _load_image_index_overrides(
        source_path, target_dir=target_dir)
    _m3_count = 0
    if _m3_overrides is not None:
        # Apply overrides by filename match
        for img in manifest_images:
            fname = img.get("filename", "")
            pg = img.get("page") or 0
            override_key = None
            # Try page-level override first
            if pg > 0 and str(pg) in _m3_overrides:
                override_key = str(pg)
            # Try filename override
            if fname in _m3_overrides:
                override_key = fname
            if override_key is not None:
                entry = _m3_overrides[override_key]
                new_cls = entry.get("classification", "").upper()
                if new_cls in ("SUB", "SUBSTANTIVE"):
                    if not img.get("is_substantive"):
                        img["is_substantive"] = True
                        img["classification_reason"] = (
                            "m3 override → SUB")
                        substantive_count += 1
                        _m3_count += 1
                elif new_cls in ("DEC", "DECORATIVE"):
                    if img.get("is_substantive"):
                        img["is_substantive"] = False
                        img["classification_reason"] = (
                            "m3 override → DEC")
                        substantive_count -= 1
                        _m3_count += 1
        if _m3_count > 0:
            print(f"  m3: Applied {_m3_count} override(s)")

    # ── 5. Write image index companion file ──
    now = datetime.now(timezone.utc)
    timestamp = now.strftime("%Y-%m-%d %H:%M")

    lines = []
    lines.append(f"# Image Index: {output_md.stem}")
    lines.append("")
    lines.append(f"Source: {source_path.resolve()}")
    lines.append(f"Converted: {output_md.resolve()}")
    lines.append(f"Generated: {timestamp}")
    lines.append(f"Pipeline version: v{PIPELINE_VERSION}")
    lines.append(f"Extractor: {extractor_label} (manifest-based index)")
    lines.append("")
    lines.append(f"Total pages: {total_pages}")
    lines.append(f"Pages with images: {pages_with_images}")
    lines.append(f"Total images detected: {total_images}")
    _fitz_fb_ct = manifest_data.get("fitz_fallback_count", 0)
    _filter_note = "(MinerU extraction)"
    if _fitz_fb_ct > 0 and _m3_count > 0:
        _filter_note = (f"(MinerU + {_fitz_fb_ct} fitz fallback + "
                        f"{_m3_count} override(s))")
    elif _fitz_fb_ct > 0:
        _filter_note = (f"(MinerU + {_fitz_fb_ct} fitz fallback)")
    elif _m3_count > 0:
        _filter_note = (f"(MinerU extraction + "
                        f"{_m3_count} override(s))")
    lines.append(f"Estimated substantive images: "
                 f"{substantive_count} {_filter_note}")
    lines.append(f"Substantive image files: {substantive_count}")
    lines.append(f"Decorative image files: "
                 f"{total_images - substantive_count}")
    lines.append("")
    lines.append("---")
    lines.append("")

    # Page-by-page index
    lines.append("## Page-by-Page Index")
    lines.append("")
    lines.append("| Page | Images | Substantive | "
                 "Context (first 150 chars) |")
    lines.append("|------|--------|-------------|"
                 "---------------------------|")

    for pg_num in sorted(page_map.keys()):
        if pg_num <= 0:
            continue
        imgs = page_map[pg_num]
        sub_on_page = any(i.get("is_substantive") for i in imgs)
        subst = "Yes" if sub_on_page else "No"
        # Try to get context from section_context
        ctx = ""
        for i in imgs:
            sc = i.get("section_context", {})
            if isinstance(sc, dict) and sc.get("heading"):
                ctx = sc["heading"][:150]
                break
        if not ctx:
            cap = next(
                (i.get("detected_caption") for i in imgs
                 if i.get("detected_caption")), None)
            ctx = (cap[:150] if cap
                   else "[MinerU extraction — no page context]")
        ctx = ctx.replace("|", "\\|")
        lines.append(
            f"| {pg_num} | {len(imgs)} "
            f"| {subst} | {ctx} |")

    if not any(pg > 0 for pg in page_map):
        lines.append(
            "| — | — | — | No images found in document |")

    lines.append("")
    lines.append("---")
    lines.append("")

    # Per-image classification detail table
    lines.append("## Per-Image Classification Detail")
    lines.append("")
    lines.append(
        "| Page | Image | Dims | File Size | Class | Reason |")
    lines.append(
        "|------|-------|------|-----------|-------|--------|")

    for img in manifest_images:
        pg = img.get("page") or 0
        fname = img.get("filename", "unknown")
        w = img.get("width", 0)
        h = img.get("height", 0)
        fsize = ""
        fp = img.get("file_path", "")
        if fp:
            try:
                fsize = "{:,}B".format(
                    _os_mineru.path.getsize(fp))
            except OSError:
                fsize = "?"
        cls = "SUB" if img.get("is_substantive") else "DEC"
        reason = img.get("classification_reason", "unknown")
        lines.append(
            "| {} | {} | {}x{} | {} | {} | {} |".format(
                pg, fname, w, h, fsize, cls, reason))

    lines.append("")
    lines.append("---")
    lines.append("")

    # Substantive images only
    lines.append("## Substantive Images Only")
    lines.append("")
    lines.append("| Page | Images | Context |")
    lines.append("|------|--------|---------|")

    sub_images = [i for i in manifest_images
                  if i.get("is_substantive")]
    # Group by page
    sub_by_page = {}
    for i in sub_images:
        pg = i.get("page") or 0
        sub_by_page.setdefault(pg, []).append(i)

    for pg_num in sorted(sub_by_page.keys()):
        if pg_num <= 0:
            continue
        imgs = sub_by_page[pg_num]
        ctx = ""
        for i in imgs:
            sc = i.get("section_context", {})
            if isinstance(sc, dict) and sc.get("heading"):
                ctx = sc["heading"][:150]
                break
        if not ctx:
            cap = next(
                (i.get("detected_caption") for i in imgs
                 if i.get("detected_caption")), None)
            ctx = cap[:150] if cap else "[MinerU extraction]"
        ctx = ctx.replace("|", "\\|")
        lines.append(f"| {pg_num} | {len(imgs)} | {ctx} |")

    if not sub_images:
        lines.append(
            "| — | — | No substantive images found |")

    lines.append("")

    # Write the file
    try:
        index_path.write_text(
            "\n".join(lines), encoding="utf-8")
        print(f"  Written: {index_path}")
        print(f"  Total images: {total_images} "
              f"({substantive_count} SUB / "
              f"{total_images - substantive_count} DEC)")
    except Exception as e:
        _warn(f"Could not write MinerU image index: {e}")
        return None

    has_testable = substantive_count > 0
    return {
        "image_index_path": str(index_path),
        "image_index_generated_at": now.isoformat(),
        "total_pages": total_pages,
        "pages_with_images": pages_with_images,
        "total_images_detected": total_images,
        "substantive_images": substantive_count,
        "has_testable_images": has_testable,
    }


def generate_image_index(source_path: Path,
                         output_md: Path,
                         fmt: str,
                         target_dir: Optional[Path] = None,
                         images_dir: Optional[Path] = None,
                         extractor: Optional[str] = None,
                         ) -> Optional[dict]:
    """R19 + m3: Generate per-file image index manifest.

    Scans the source document for images, classifies decorative vs
    substantive, applies m3 overrides if an image-index-overrides.json
    file is found, and writes a structured markdown manifest alongside
    the converted .md file.

    When extractor is "mineru", the image index is generated from the
    MinerU manifest (image-manifest.json) rather than scanning the
    original PDF with fitz.  This ensures the image index matches the
    actual images MinerU extracted, avoiding count mismatches with the
    fitz xref scan that would cause qc-structural.py consistency
    check failures.

    Args:
        source_path: Path to the source document (PDF/PPTX/DOCX).
        output_md: Path to the converted .md file.
        fmt: File format string ("pdf", "pptx", "docx").
        target_dir: Optional target directory (--target-dir).
                    Used as fallback location for override file.
        images_dir: Optional path to the images directory where
                    convert-paper.py saves extracted raster images.
                    Vector renders are saved here so that file
                    organization (Step 9b) picks them up correctly.
                    BUG-1 fix: without this, vector renders were
                    saved to images/ root instead of the slug
                    subdirectory (images/<slug>/).
        extractor: Optional extractor name (e.g. "mineru", "pymupdf4llm").
                   When "mineru", reads the MinerU manifest instead of
                   scanning the PDF with fitz.

    Returns:
        Dict with image index metadata (for R21 registry integration),
        or None if scanning failed. Keys:
            image_index_path, total_pages, pages_with_images,
            total_images_detected, substantive_images, has_testable_images
    """
    print(f"\n{'─' * 40}")
    print("Step 6c: Image Index Generation (R19)")
    print("─" * 40)

    # Determine output path for the image index
    index_stem = output_md.stem
    index_path = output_md.parent / f"{index_stem}-image-index.md"

    # ── MinerU path: generate image index from MinerU manifest ──
    # When MinerU was the extractor, the manifest written by
    # _normalize_mineru_output() is the authoritative source of
    # image data.  Scanning the PDF with fitz would produce
    # different counts (fitz counts xrefs, MinerU extracts
    # differently), causing qc-structural.py consistency failures.
    if extractor == "mineru" and fmt == "pdf":
        return _generate_image_index_from_mineru_manifest(
            source_path=source_path,
            output_md=output_md,
            index_path=index_path,
            images_dir=images_dir,
            target_dir=target_dir,
        )

    # ── RC-B: pymupdf4llm manifest-based index ──────────────────────────
    # When using the pymupdf4llm extractor, prefer image-manifest.json as
    # the source of truth over a fresh fitz xref scan.  The xref scan
    # misses split-panel images (fig{N}a/fig{N}b files) created by
    # _split_panels() in convert-paper.py, producing ghost xref entries
    # and an index that under-counts disk files (e.g. 151 vs 211 files).
    # image-manifest.json is written by convert-paper.py AFTER panel
    # splitting, so it always matches the actual files on disk.
    if fmt == "pdf" and extractor != "mineru":
        _rcb_manifest_path = None
        if images_dir is not None:
            _rcb_candidate = images_dir / "image-manifest.json"
            if _rcb_candidate.exists():
                _rcb_manifest_path = _rcb_candidate
        if _rcb_manifest_path is None:
            # Fallback: check common manifest locations
            for _rcb_dir in [
                output_md.parent / "images",
                output_md.parent / f"{output_md.stem}_images",
            ]:
                _rcb_try = _rcb_dir / "image-manifest.json"
                if _rcb_try.exists():
                    _rcb_manifest_path = _rcb_try
                    break
        if _rcb_manifest_path is not None:
            _rcb_label = extractor if extractor else "pymupdf4llm"
            print(f"  RC-B: image-manifest.json found — using manifest "
                  f"(not xref scan) for {_rcb_label} extractor")
            return _generate_image_index_from_mineru_manifest(
                source_path=source_path,
                output_md=output_md,
                index_path=index_path,
                images_dir=images_dir,
                target_dir=target_dir,
                extractor_label=_rcb_label,
            )

    # Scan based on format
    if fmt == "pdf":
        pages = scan_pdf_images(str(source_path))
    else:
        # PPTX/DOCX scanning is handled by convert-office.py agent
        # For now, only PDF is automated in run-pipeline.py
        print(f"  SKIP: Image index for {fmt.upper()} is handled by "
              "convert-office.py")
        return None

    if pages is None:
        # Encrypted or unreadable
        _write_error_image_index(index_path, source_path,
                                 "Could not scan: encrypted or unreadable document")
        print(f"  ⚠ Could not scan document (encrypted/unreadable)")
        print(f"  Written: {index_path}")
        return {
            "image_index_path": str(index_path),
            "image_index_generated_at": datetime.now(timezone.utc).isoformat(),
            "total_pages": 0,
            "pages_with_images": 0,
            "total_images_detected": 0,
            "substantive_images": 0,
            "has_testable_images": False,
        }

    # ── FIX-1: Render pure-vector pages as PNG images ──────────────────
    # Pages with 0 raster images but significant vector content (e.g.
    # Kaplan-Meier curves, flowcharts, bar charts) would otherwise be
    # invisible to the pipeline.  Render them at 300 DPI and add to
    # the images directory and manifest.
    vector_rendered = []  # Track rendered vector pages for manifest
    if fmt == "pdf" and _HAS_FITZ:
        # BUG-1 fix: use the caller-provided images_dir (same directory
        # where convert-paper.py saves extracted raster images, typically
        # images/<slug>/).  This ensures vector renders are found by Step
        # 9b file organization which looks for images in slug subdirs.
        # Fallback chain: caller images_dir → images/ root → stem_images
        # → caller images_dir (create if needed).
        if images_dir is not None and images_dir.exists():
            images_dir_for_vectors = images_dir
        else:
            images_dir_for_vectors = output_md.parent / "images"
            if not images_dir_for_vectors.exists():
                # Try stem-based images dir pattern
                _stem_images = output_md.parent / f"{output_md.stem}_images"
                if _stem_images.exists():
                    images_dir_for_vectors = _stem_images
                else:
                    # Use caller-provided path even if it doesn't exist yet;
                    # create it so renders have a home in the right place.
                    if images_dir is not None:
                        images_dir_for_vectors = images_dir
                        images_dir_for_vectors.mkdir(parents=True, exist_ok=True)
                    else:
                        # No slug dir known and no existing dirs found.
                        # Do NOT create images/ root — vector renders there
                        # would be missed by Step 9b file organization.
                        _warn("Vector renders skipped: no valid images "
                              "directory found and images_dir not provided.")
                        images_dir_for_vectors = None

        if images_dir_for_vectors is not None:
            try:
                doc = fitz.open(str(source_path))
                for page_data in pages:
                    if (page_data["image_count"] == 0
                            and _has_significant_vector_content(page_data)):
                        page_num = page_data["page"] - 1  # 0-indexed
                        try:
                            page = doc[page_num]
                            pix = page.get_pixmap(dpi=300)
                            render_name = (f"page{page_data['page']}"
                                           f"-vector-render.png")
                            render_path = images_dir_for_vectors / render_name
                            pix.save(str(render_path))
                            # MINOR-1: capture dimensions then release pixmap
                            # to avoid memory accumulation on PDFs with many
                            # vector pages.
                            _pix_w, _pix_h = pix.width, pix.height
                            del pix
                            if _ensure_max_dimension(render_path):
                                # Re-read dimensions after resize
                                try:
                                    from PIL import Image as _PILDim
                                    with _PILDim.open(render_path) as _resized:
                                        _pix_w, _pix_h = _resized.size
                                except Exception:
                                    pass
                            # Update page data to reflect the rendered image
                            page_data["image_count"] = 1
                            page_data["image_details"] = [{
                                "xref": None,
                                "width": _pix_w,
                                "height": _pix_h,
                                "source": "vector-render",
                            }]
                            page_data["vector_rendered"] = True
                            vector_rendered.append({
                                "page": page_data["page"],
                                "filename": render_name,
                                "file_path": str(render_path),
                                "width": _pix_w,
                                "height": _pix_h,
                                "source": "vector-render",
                                "drawing_count": page_data.get(
                                    "drawing_count", 0),
                            })
                            print(f"  FIX-1: Rendered vector page "
                                  f"{page_data['page']} → {render_name} "
                                  f"({_pix_w}x{_pix_h}px)")
                        except Exception as e:
                            _warn(f"Could not render vector page "
                                  f"{page_data['page']}: {e}")
                doc.close()
            except Exception as e:
                _warn(f"Could not open PDF for vector rendering: {e}")

            # Update manifest JSON with vector-rendered images
            if vector_rendered:
                _update_manifest_with_vector_renders(
                    output_md, vector_rendered, images_dir_for_vectors)

                # FIX-1.7: Update YAML image_notes from "none" to "pending"
                # when vector renders were added. Step 3 sets image_notes
                # based on raster count (which may be 0), but vector renders
                # discovered here mean images DO exist for analysis.
                try:
                    _yaml_content = output_md.read_text(encoding="utf-8")
                    if "image_notes: none" in _yaml_content:
                        _yaml_content = _yaml_content.replace(
                            "image_notes: none",
                            "image_notes: pending",
                            1,  # replace only the first occurrence (YAML header)
                        )
                        output_md.write_text(_yaml_content, encoding="utf-8")
                        print(f"  FIX-1.7: Updated YAML image_notes: "
                              f"none → pending ({len(vector_rendered)} "
                              f"vector render(s))")
                except Exception as _yaml_e:
                    _warn(f"Could not update YAML image_notes "
                          f"after vector renders: {_yaml_e}")

    # Build xref-to-filepath map for per-image classification.
    # Must run AFTER FIX-1 (vector renders need to be in images_dir)
    # and BEFORE classification loop.
    if fmt == "pdf":
        _build_xref_filepath_map(images_dir, pages)

    total_pages = len(pages)
    pages_with_images = sum(1 for p in pages if p["image_count"] > 0)
    total_images = sum(p["image_count"] for p in pages)

    # Pre-build xref frequency map once (avoids O(P*I) rebuild per page)
    _xref_counts = {}
    for _p in pages:
        for _img in _p.get("image_details", []):
            _xr = _img.get("xref")
            if _xr is not None:
                _xref_counts[_xr] = _xref_counts.get(_xr, 0) + 1

    # Classify each page
    for page_data in pages:
        if page_data["image_count"] > 0:
            page_data["is_substantive"] = _classify_page_images(
                page_data, pages, total_pages,
                xref_counts=_xref_counts)
        elif _has_significant_vector_content(page_data):
            # Pure-vector page that wasn't rendered (non-PDF or fitz
            # unavailable).  Still classify via heuristic 8 + area check.
            page_data["is_substantive"] = _classify_page_images(
                page_data, pages, total_pages,
                xref_counts=_xref_counts)
        else:
            page_data["is_substantive"] = False

    # ── FIX-2: Per-file decorative detection (PDF only) ─────────────
    # Check extracted image files for blankness (all-white, near-uniform
    # color, tiny file size) AND other decorative signals (color blocks,
    # dark covers, small file size relative to dimensions).
    # M1/M2: enhanced to also detect color blocks and dark cover patterns.
    # Reclassify pages as decorative only when ALL images on that page
    # are decorative (conservative: one substantive image keeps the page).
    if fmt == "pdf":
        _decorative_count = 0
        # Find images directory for file-level checks.
        # BUG-1 fix: Use caller-provided images_dir (slug subdirectory)
        # first.  The old code searched images/ root which contains only
        # subdirectories, so glob found zero image files.
        _imgs_dir = None
        if images_dir is not None and images_dir.exists():
            _imgs_dir = images_dir
        else:
            _imgs_dir_candidates = [
                output_md.parent / f"{output_md.stem}_images",
                output_md.parent / "images",
            ]
            _imgs_dir = next(
                (d for d in _imgs_dir_candidates if d.exists()), None)
        # Track decorative files per page for page-level reclassification
        _page_decorative: dict = {}  # page_num -> set of decorative filenames
        _page_total: dict = {}       # page_num -> set of all filenames

        # S36-FIX: MAJOR-02 - Two-pass approach. First pass builds complete
        # _page_total so that ALL files on a page are counted before any
        # classification decisions use page_image_count. Without this, early
        # files on a page see an incomplete count (e.g. first file sees 1,
        # second sees 2) causing inconsistent branding decisions.
        _img_file_list = []  # (Path, page_num_or_None) pairs
        if _imgs_dir is not None:
            for _img_file in sorted(_imgs_dir.iterdir()):
                if not _img_file.is_file():
                    continue
                if _img_file.suffix.lower() not in (
                        ".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".tif",
                        ".gif"):
                    continue
                # Extract page number from filename.
                # m-4 FIX: Try multiple patterns to handle different
                # extraction tool naming conventions:
                #   pymupdf: page6-img1.png, page6-vector-render.png
                #   fallback: p6_img1.png, img-page6.png
                # Images without any page number are still checked for
                # decorative status (below) but cannot participate in
                # page-level reclassification.
                _pg_match = re.search(
                    r'page[\-_]?(\d+)', _img_file.stem, re.IGNORECASE)
                if not _pg_match:
                    # Fallback: try p<N> pattern (e.g. p6_img1.png)
                    _pg_match = re.search(
                        r'(?:^|[\-_])p(\d+)(?:[\-_]|$)',
                        _img_file.stem, re.IGNORECASE)
                _pg_num = int(_pg_match.group(1)) if _pg_match else None
                if _pg_num is not None:
                    _page_total.setdefault(_pg_num, set()).add(
                        _img_file.name)
                _img_file_list.append((_img_file, _pg_num))

            # S36-FIX: MAJOR-02 - Second pass: classify with complete counts
            for _img_file, _pg_num in _img_file_list:
                _is_dec = False
                # Check 1: blank/color-block/dark-cover via enhanced
                # _is_blank_image (includes M1/M2 heuristics)
                if _is_blank_image(str(_img_file)):
                    _is_dec = True
                # Check 2: journal branding via file size + dimensions.
                # _is_journal_branding requires width/height; read from
                # PIL if available.  This catches small file-size images
                # (< 5KB) that are not caught by _is_blank_image.
                # S36: Pass page_num and page_image_count for enhanced checks.
                if not _is_dec:
                    try:
                        import os as _os_f2
                        _f2_size = _os_f2.path.getsize(str(_img_file))
                        _f2_w, _f2_h = 0, 0
                        try:
                            from PIL import Image as _PILf2
                            _f2_img = _PILf2.open(str(_img_file))
                            _f2_w, _f2_h = _f2_img.size
                            _f2_img.close()
                        except Exception:
                            pass
                        # S36-FIX: MAJOR-02 - Use complete page count
                        _f2_pg_img_count = (
                            len(_page_total.get(_pg_num, set()))
                            if _pg_num is not None else 0
                        )
                        if _is_journal_branding(
                                {"width": _f2_w, "height": _f2_h},
                                file_size_bytes=_f2_size,
                                page_num=_pg_num or 0,
                                page_image_count=_f2_pg_img_count):
                            _is_dec = True
                    except Exception:
                        pass

                if _is_dec:
                    _decorative_count += 1
                    if _pg_num is not None:
                        _page_decorative.setdefault(_pg_num, set()).add(
                            _img_file.name)
                    # m-4 FIX: warn about orphan decorative images that
                    # cannot trigger page-level reclassification
                    _orphan_tag = ""
                    if _pg_num is None:
                        _orphan_tag = " (no page number — skipped " \
                                      "page reclassification)"
                    print(f"  FIX-2: Decorative image detected: "
                          f"{_img_file.name}{_orphan_tag}")

        # Per-image reclassification: mark individual images as DEC
        # when FIX-2 file scan detects them as decorative.  Then
        # recalculate page verdict from the per-image flags.
        for _pg_num, _dec_files in _page_decorative.items():
            for _pd in pages:
                if _pd["page"] == _pg_num:
                    # Try file_path matching first
                    _matched_any = False
                    for img in _pd.get("image_details", []):
                        fp = img.get("file_path", "")
                        if fp and Path(fp).name in _dec_files:
                            img["is_substantive"] = False
                            if not img.get("classification_reason"):
                                img["classification_reason"] = "fix2_file_scan"
                            _matched_any = True
                    # Fallback: when file_path is not populated for ANY
                    # image on this page, file_path matching cannot work.
                    # In that case, if ALL files for this page are
                    # decorative, mark all images on the page as DEC.
                    if not _matched_any:
                        _all_page_files = _page_total.get(_pg_num, set())
                        if (_all_page_files
                                and _all_page_files.issubset(_dec_files)):
                            for img in _pd.get("image_details", []):
                                img["is_substantive"] = False
                                if not img.get("classification_reason"):
                                    img["classification_reason"] = \
                                        "fix2_file_scan"
                    # Recalculate page verdict from per-image flags
                    _any_sub = any(
                        img.get("is_substantive", False)
                        for img in _pd.get("image_details", []))
                    if not _any_sub:
                        _pd["is_substantive"] = False
                        _pd["blank_detected"] = True
                    break

        if _decorative_count > 0:
            print(f"  FIX-2: {_decorative_count} decorative image(s) "
                  f"detected (file-level scan)")

    # ── m3: Apply image-index-overrides.json (if present) ──
    # Overrides are applied AFTER automatic heuristic classification
    # but BEFORE the manifest is written.  Modifies pages in-place.
    _m3_override_count = 0
    _m3_overrides = _load_image_index_overrides(
        source_path, target_dir=target_dir)
    if _m3_overrides is not None:
        _m3_override_count = _apply_overrides_to_pages(
            pages, _m3_overrides, source_path.name)
        if _m3_override_count > 0:
            print(f"  m3: Applied {_m3_override_count} override(s)")

    substantive_pages = [p for p in pages
                         if p["is_substantive"]]
    decorative_pages = [p for p in pages
                        if p["image_count"] > 0 and not p["is_substantive"]]

    substantive_count = len(substantive_pages)

    # Write the manifest
    # R21: Use UTC for image_index_generated_at (ISO 8601 consistency)
    now = datetime.now(timezone.utc)
    timestamp = now.strftime("%Y-%m-%d %H:%M")

    lines = []
    lines.append(f"# Image Index: {output_md.stem}")
    lines.append("")
    lines.append(f"Source: {source_path.resolve()}")
    lines.append(f"Converted: {output_md.resolve()}")
    lines.append(f"Generated: {timestamp}")
    lines.append(f"Pipeline version: v{PIPELINE_VERSION}")
    lines.append("")
    lines.append(f"Total pages: {total_pages}")
    lines.append(f"Pages with images: {pages_with_images}")
    lines.append(f"Total images detected: {total_images}")
    # m3: Annotate the count when overrides were applied
    _filter_note = "(after filtering)"
    if _m3_override_count > 0:
        _filter_note = (f"(after filtering + "
                        f"{_m3_override_count} override(s))")
    lines.append(f"Estimated substantive images: {substantive_count} "
                 f"{_filter_note}")
    # Per-image SUB/DEC counts
    # Default True: unclassified images are conservatively treated as SUB
    _per_image_sub = sum(
        1 for p in pages for img in p.get("image_details", [])
        if img.get("is_substantive", True)
    )
    _per_image_dec = sum(
        1 for p in pages for img in p.get("image_details", [])
        if not img.get("is_substantive", True)
    )
    lines.append(f"Substantive image files: {_per_image_sub}")
    lines.append(f"Decorative image files: {_per_image_dec}")
    lines.append("")
    lines.append("---")
    lines.append("")

    # Page-by-page index
    lines.append("## Page-by-Page Index")
    lines.append("")
    lines.append("| Page | Images | Substantive | "
                 "Context (first 150 chars) |")
    lines.append("|------|--------|-------------|"
                 "---------------------------|")

    for page_data in pages:
        if page_data["image_count"] > 0:
            subst = "Yes" if page_data["is_substantive"] else "No"
            ctx = page_data["context"] or "[no text on page]"
            # Escape pipe characters in context for markdown table
            ctx = ctx.replace("|", "\\|")
            lines.append(
                f"| {page_data['page']} | {page_data['image_count']} "
                f"| {subst} | {ctx} |"
            )

    if pages_with_images == 0:
        lines.append("| — | — | — | No images found in document |")

    lines.append("")
    lines.append("---")
    lines.append("")

    # Per-image classification detail table
    import os as _os_idx
    lines.append("## Per-Image Classification Detail")
    lines.append("")
    lines.append("| Page | Image | Dims | File Size | Class | Reason |")
    lines.append("|------|-------|------|-----------|-------|--------|")

    for page_data in pages:
        for img in page_data.get("image_details", []):
            fp = img.get("file_path", "")
            fname = Path(fp).name if fp else "xref-%s" % (
                img.get("xref", "?"),)
            w = img.get("width", 0)
            h = img.get("height", 0)
            _fsize = ""
            if fp:
                try:
                    _fsize = "{:,}B".format(_os_idx.path.getsize(fp))
                except OSError:
                    _fsize = "?"
            cls = "SUB" if img.get("is_substantive", True) else "DEC"
            reason = img.get("classification_reason", "unknown")
            lines.append(
                "| {} | {} | {}x{} | {} | {} | {} |".format(
                    page_data["page"], fname, w, h,
                    _fsize, cls, reason)
            )

    lines.append("")
    lines.append("---")
    lines.append("")

    # Substantive images only
    lines.append("## Substantive Images Only")
    lines.append("")
    lines.append("| Page | Images | Context |")
    lines.append("|------|--------|---------|")

    for page_data in substantive_pages:
        ctx = page_data["context"] or "[no text on page]"
        ctx = ctx.replace("|", "\\|")
        # Use per-image SUB count instead of total image count
        _sub_count = sum(
            1 for img in page_data.get("image_details", [])
            if img.get("is_substantive", False))
        lines.append(
            f"| {page_data['page']} | {_sub_count} "
            f"| {ctx} |"
        )

    if not substantive_pages:
        lines.append("| — | — | No substantive images found |")

    lines.append("")
    lines.append("---")
    lines.append("")

    # ── FIX-4: Vector Content Notes ────────────────────────────────
    # Warn about pages with BOTH raster images AND significant vector
    # content (the raster images may not capture vector elements).
    # Also note pages that were vector-rendered by FIX-1.
    _vector_notes = []
    for page_data in pages:
        dc = page_data.get("drawing_count", 0)
        ic = page_data.get("image_count", 0)
        if _has_significant_vector_content(page_data):
            if page_data.get("vector_rendered"):
                _vector_notes.append(
                    f"- Page {page_data['page']}: {dc} vector "
                    f"drawings detected (page rendered as PNG)")
            elif ic > 0:
                _vector_notes.append(
                    f"- Page {page_data['page']}: {dc} vector "
                    f"drawings detected (figure may contain "
                    f"additional vector elements)")

    if _vector_notes:
        lines.append("## Vector Content Notes")
        lines.append("")
        for note in _vector_notes:
            lines.append(note)
        lines.append("")
        lines.append("---")
        lines.append("")

    # ── FIX-5: Image file count vs page count ──────────────────────
    # Count actual image files on disk to report alongside page count.
    # Panel splits (e.g. fig1a, fig1b) may cause file count > page count.
    _image_file_count = 0
    _imgs_dir_candidates_f5 = []
    if images_dir is not None:
        _imgs_dir_candidates_f5.append(images_dir)
    _imgs_dir_candidates_f5.extend([
        output_md.parent / "images",
        output_md.parent / f"{output_md.stem}_images",
    ])
    _imgs_dir_f5 = next(
        (d for d in _imgs_dir_candidates_f5 if d.exists()), None)
    if _imgs_dir_f5 is not None:
        _image_exts = {".png", ".jpg", ".jpeg", ".bmp",
                       ".tiff", ".tif", ".gif", ".svg"}
        _image_file_count = sum(
            1 for f in _imgs_dir_f5.iterdir()
            if f.is_file() and f.suffix.lower() in _image_exts)

    # Filtering summary
    lines.append("## Filtering Summary")
    lines.append("")
    lines.append(f"- Pages scanned: {total_pages}")
    lines.append(f"- Pages with images: {pages_with_images}")
    # FIX-5: Show both page count and file count when they differ
    if _image_file_count > 0 and _image_file_count != total_images:
        lines.append(f"- Total image files on disk: {_image_file_count}")
    lines.append(f"- Images classified SUB: {_per_image_sub}")
    lines.append(f"- Images classified DEC: {_per_image_dec}")
    lines.append(f"- Pages classified as decorative: {len(decorative_pages)}")
    lines.append(f"- Pages classified as substantive: {substantive_count}")
    # FIX-1: Note vector renders in criteria list
    _vr_count = len(vector_rendered) if fmt == "pdf" else 0
    _criteria = ("dimensions (<50x50px), "
                 "repeated xref (>50% pages), title/cover page, "
                 "last-page decorations, figure keywords, "
                 "journal branding (<100px/<5KB/extreme aspect), "
                 "blank image detection (3-tier: size<2KB/near-black/<16colors/std<5), "
                 "M1 color-block (<32 unique colors), "
                 "low-density badge (<0.15 B/px AND <50 colors), "
                 "low word count render (<20 words)")
    if _vr_count > 0:
        _criteria += f", vector rendering ({_vr_count} pages)"
    lines.append(f"- Filtering criteria applied: {_criteria}")
    # m3: Note overrides in the summary when applied
    if _m3_override_count > 0:
        lines.append(f"- Manual overrides applied: {_m3_override_count} "
                     f"page(s) (via image-index-overrides.json)")

    # Write to disk
    try:
        index_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
        print(f"  GENERATED: {index_path.name}")
        print(f"    Pages scanned: {total_pages}")
        print(f"    Pages with images: {pages_with_images}")
        print(f"    Total images: {total_images}")
        pct = (substantive_count / pages_with_images * 100
               if pages_with_images > 0 else 0)
        print(f"    Substantive: {substantive_count} ({pct:.0f}%)")
        print(f"    Decorative filtered: {len(decorative_pages)} "
              f"({100 - pct:.0f}%)")
        print(f"    Per-image SUB: {_per_image_sub}, DEC: {_per_image_dec}")
    except Exception as e:
        _warn(f"Could not write image index: {e}")
        return None

    # Write per-image classification to manifest JSON for
    # prepare-image-analysis.py integration.
    _manifest_path = None
    if images_dir is not None:
        _manifest_path = images_dir / "image-manifest.json"
        if not _manifest_path.exists():
            # Try alternative location: same dir as output .md
            _alt_manifest = output_md.parent / (
                output_md.stem.replace("-converted", "") + "_manifest.json")
            if _alt_manifest.exists():
                _manifest_path = _alt_manifest
    if _manifest_path is None or not _manifest_path.exists():
        # Try common manifest locations
        for _mdir in [output_md.parent / "images",
                      output_md.parent / (output_md.stem + "_images")]:
            _try = _mdir / "image-manifest.json"
            if _try.exists():
                _manifest_path = _try
                break

    if _manifest_path is not None and _manifest_path.exists():
        try:
            _manifest = json.loads(
                _manifest_path.read_text(encoding="utf-8"))
            # Build per-image classification map: filename -> {class, reason}
            _img_class_map = {}
            for _page_data in pages:
                for img in _page_data.get("image_details", []):
                    fp = img.get("file_path", "")
                    if fp:
                        fname = Path(fp).name
                        _img_class_map[fname] = {
                            "classification": (
                                "SUB" if img.get("is_substantive", True)
                                else "DEC"),
                            "reason": img.get(
                                "classification_reason", "unknown"),
                        }
            _manifest["per_image_classification"] = _img_class_map
            _manifest_path.write_text(
                json.dumps(_manifest, indent=2, ensure_ascii=False) + "\n",
                encoding="utf-8")
            print(f"  Manifest updated with per-image classification "
                  f"({len(_img_class_map)} images)")
        except Exception as e:
            _warn("Could not update manifest with per-image "
                  "classification: %s" % e)

    return {
        "image_index_path": str(index_path),
        "image_index_generated_at": now.isoformat(),
        "total_pages": total_pages,
        "pages_with_images": pages_with_images,
        "total_images_detected": total_images,
        "substantive_images": substantive_count,
        "has_testable_images": substantive_count > 0,
    }


def _write_error_image_index(index_path: Path, source_path: Path,
                              error_msg: str) -> None:
    """Write an image index manifest with an error note."""
    now = datetime.now(timezone.utc)
    lines = [
        f"# Image Index: {index_path.stem.replace('-image-index', '')}",
        "",
        f"Source: {source_path.resolve()}",
        f"Generated: {now.strftime('%Y-%m-%d %H:%M')}",
        f"Pipeline version: v{PIPELINE_VERSION}",
        "",
        f"**Error:** {error_msg}",
        "",
        "---",
        "",
        "## Page-by-Page Index",
        "",
        "| Page | Images | Substantive | Context (first 150 chars) |",
        "|------|--------|-------------|---------------------------|",
        f"| — | — | — | {error_msg} |",
    ]
    try:
        index_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    except Exception as e:
        _warn(f"Could not write error image index: {e}")


# ═══════════════════════════════════════════════════════════════════════════
# R20: TESTABLE IMAGE INDEX (Project-Level Aggregate)
# ═══════════════════════════════════════════════════════════════════════════

# Default HTA topic keywords (used when no project config exists)
_DEFAULT_HTA_TOPICS = {
    "Introduction to HTA": [
        "introduction", "framework", "economic evaluation", "hta"],
    "Costs & Costing": [
        "cost", "costing", "resource use", "friction", "human capital"],
    "Quality of Life / HRQoL": [
        "quality of life", "hrqol", "eq-5d", "sf-6d", "utility", "qaly"],
    "Modelling": [
        "model", "markov", "decision tree", "simulation", "extrapolation"],
    "Cost-Effectiveness Analysis": [
        "icer", "ce plane", "threshold", "cost-effectiveness", "nmb"],
    "Sensitivity & Uncertainty": [
        "sensitivity", "uncertainty", "tornado", "psa", "ceac", "bootstrap"],
    "Discounting": [
        "discount", "time preference", "present value"],
    "HTA & Policy Making": [
        "policy", "reimbursement", "value flower", "nice", "zin"],
    "Transferability": [
        "transferability", "generalisability", "country", "jurisdiction"],
    "Equity & Distribution": [
        "equity", "distribution", "qaly weight", "fair innings"],
    "Theoretical Foundations": [
        "welfare", "welfarism", "extra-welfarism", "capability"],
}

_DEFAULT_SOURCE_PATTERNS = {
    "CURRENT SLIDES": ["current", "2026"],
    "PREVIOUS SLIDES": ["previous", "2025"],
    "LITERATURE": ["literature", "paper", "article"],
    "WORKING GROUP": ["wg", "working group", "workgroup"],
}


def _load_topic_config(project_dir: Path) -> tuple:
    """Load project-specific topic config if available.

    Returns (topics_dict_or_None, source_patterns_dict).
    When no config file is found, returns (None, _DEFAULT_SOURCE_PATTERNS).
    A None topics value signals that the caller should use generic
    by-document grouping (group by source filename, no topic classification),
    as specified by the R20 requirements.
    """
    config_path = project_dir / ".claude" / "config" / "image-index-topics.json"

    if config_path.exists():
        try:
            data = json.loads(config_path.read_text(encoding="utf-8"))
            topics = data.get("topics")
            if topics is None:
                _warn(f"Topic config at {config_path} has no 'topics' key. "
                      "Using generic by-document grouping.")
            source_patterns = data.get("source_patterns",
                                       _DEFAULT_SOURCE_PATTERNS)
            return (topics, source_patterns)
        except (json.JSONDecodeError, KeyError) as e:
            _warn(f"Could not parse topic config at {config_path}: {e}. "
                  "Using generic by-document grouping.")

    # No config file: generic by-document grouping (R20 spec)
    return (None, _DEFAULT_SOURCE_PATTERNS)


def _classify_topic(context: str, filename: str,
                    topics: Optional[dict]) -> str:
    """Classify an image entry into a topic using weighted keyword matching.

    When topics is None (no project config), returns the filename stem as
    the group name (generic by-document grouping per R20 spec).

    Algorithm (from requirements, when topics dict is provided):
      1. Combine lowercase(context) + lowercase(filename)
      2. For each topic: count keyword matches
         - exact match = 2, substring match = 1
      3. Assign to highest-scoring topic
      4. If no match (score=0): "Uncategorised"
      5. If tied: first topic in config order
    """
    # Generic by-document grouping when no topic config exists
    if topics is None:
        stem = Path(filename).stem if filename else "Unknown Document"
        return stem
    combined = (context + " " + filename).lower()
    best_topic = "Uncategorised"
    best_score = 0

    for topic_name, keywords in topics.items():
        score = 0
        for kw in keywords:
            kw_lower = kw.lower()
            # Check for exact word boundary match first
            # Simple heuristic: surrounded by non-alpha chars or at boundaries
            if re.search(r'\b' + re.escape(kw_lower) + r'\b', combined):
                score += 2
            elif kw_lower in combined:
                score += 1
        if score > best_score:
            best_score = score
            best_topic = topic_name

    return best_topic


def _detect_source_category(file_path: str,
                             source_patterns: dict) -> str:
    """Detect source category from file path heuristics.

    Returns one of: CURRENT SLIDES, PREVIOUS SLIDES, LITERATURE,
    WORKING GROUP, or default LITERATURE.
    """
    path_lower = file_path.lower()

    # Check each category's patterns
    for category, patterns in source_patterns.items():
        for pattern in patterns:
            if pattern.lower() in path_lower:
                return category

    return "LITERATURE"  # Default


_SOURCE_PRIORITY = {
    "CURRENT SLIDES": 0,
    "PREVIOUS SLIDES": 1,
    "LITERATURE": 2,
    "WORKING GROUP": 3,
}


def _parse_image_index_file(index_path: Path) -> list:
    """Parse a per-file image index manifest and extract substantive entries.

    Returns list of dicts with keys:
        page, image_count, context, source_file, source_path
    """
    entries = []
    try:
        content = index_path.read_text(encoding="utf-8")
    except Exception as e:
        _warn(f"Could not read image index {index_path}: {e}")
        return []

    # Extract source path from header
    source_path = ""
    for line in content.splitlines():
        if line.startswith("Source: "):
            source_path = line[len("Source: "):].strip()
            break

    # Find the "Substantive Images Only" section and parse its table
    in_substantive = False
    past_header = False

    for line in content.splitlines():
        if "## Substantive Images Only" in line:
            in_substantive = True
            continue
        if in_substantive and line.startswith("|---"):
            past_header = True
            continue
        if in_substantive and line.startswith("---"):
            break  # End of section
        if in_substantive and past_header and line.startswith("|"):
            parts = [p.strip() for p in line.split("|")]
            # Filter empty parts from leading/trailing pipes
            parts = [p for p in parts if p]
            if len(parts) >= 3 and parts[0] != "—":
                try:
                    page = int(parts[0])
                    img_count = int(parts[1])
                    ctx = parts[2] if len(parts) > 2 else ""
                    entries.append({
                        "page": page,
                        "image_count": img_count,
                        "context": ctx,
                        "source_file": Path(source_path).name
                                       if source_path else index_path.stem,
                        "source_path": source_path,
                    })
                except (ValueError, IndexError):
                    continue

    return entries


# ═══════════════════════════════════════════════════════════════════════════
# m2: AGENT DESCRIPTIONS FILE GENERATION
# ═══════════════════════════════════════════════════════════════════════════

def _find_analysis_manifest(output_md: Path,
                            images_dir: Path,
                            fmt: str,
                            short_name: Optional[str] = None,
                            input_stem: Optional[str] = None,
                            ) -> Optional[Path]:
    """Locate the analysis-manifest.json for a converted document.

    Searches multiple candidate locations because prepare-image-analysis.py
    writes analysis-manifest.json into the manifest's images_dir, which
    differs between PDF and office formats.

    Args:
        output_md: Path to the converted .md file.
        images_dir: Primary images directory (from pipeline).
        fmt: File format ("pdf", "pptx", "docx").
        short_name: Optional short name used for naming.
        input_stem: Optional input file stem (fallback for naming).

    Returns:
        Path to analysis-manifest.json if found, else None.
    """
    candidates = [images_dir / "analysis-manifest.json"]

    # For office formats, the images dir may be named differently
    if fmt != "pdf":
        if short_name:
            office_images_dir = output_md.parent / f"{short_name}_images"
            candidates.append(office_images_dir / "analysis-manifest.json")
        if input_stem:
            stem_images_dir = output_md.parent / f"{input_stem}_images"
            candidates.append(stem_images_dir / "analysis-manifest.json")
        # Also try the output_md stem
        md_stem_images_dir = output_md.parent / f"{output_md.stem}_images"
        candidates.append(md_stem_images_dir / "analysis-manifest.json")

    # Deduplicate while preserving order
    seen = set()
    unique_candidates = []
    for c in candidates:
        resolved = c.resolve()
        if resolved not in seen:
            seen.add(resolved)
            unique_candidates.append(c)

    return next((p for p in unique_candidates if p.exists()), None)


def _find_image_manifest(output_md: Path,
                         images_dir: Path,
                         fmt: str,
                         short_name: Optional[str] = None,
                         input_stem: Optional[str] = None,
                         ) -> Optional[Path]:
    """Locate the image-manifest.json for a converted document.

    Args:
        output_md: Path to the converted .md file.
        images_dir: Primary images directory.
        fmt: File format ("pdf", "pptx", "docx").
        short_name: Optional short name.
        input_stem: Optional input file stem.

    Returns:
        Path to image-manifest.json if found, else None.
    """
    candidates = [images_dir / "image-manifest.json"]

    # Office formats: manifest may be at output_md.parent level
    if fmt != "pdf":
        if short_name:
            candidates.append(
                output_md.parent / f"{short_name}_manifest.json")
        if input_stem:
            candidates.append(
                output_md.parent / f"{input_stem}_manifest.json")
        candidates.append(
            output_md.parent / f"{output_md.stem}_manifest.json")

    seen = set()
    unique_candidates = []
    for c in candidates:
        resolved = c.resolve()
        if resolved not in seen:
            seen.add(resolved)
            unique_candidates.append(c)

    return next((p for p in unique_candidates if p.exists()), None)


def generate_agent_descriptions_file(
        output_md: Path,
        image_index_path: Path,
        images_dir: Path,
        fmt: str,
        short_name: Optional[str] = None,
        input_stem: Optional[str] = None,
) -> Optional[Path]:
    """m2: Generate an agent prompt file for image descriptions.

    Creates a structured file that provides a Claude subagent with
    everything needed to describe substantive images using the
    generate-image-notes.md skill template.  The file includes:
    - Paths to the markdown file, images directory, and manifests
    - The document domain / title from the analysis manifest
    - A table of all substantive images with page numbers,
      nearby text context, section headings, and file paths
    - A ready-to-use Task() invocation block

    This file is written alongside the image index so a user or
    orchestrator can feed it directly to a subagent.

    Args:
        output_md: Path to the converted .md file.
        image_index_path: Path to the *-image-index.md file.
        images_dir: Primary images directory.
        fmt: File format ("pdf", "pptx", "docx").
        short_name: Optional short name.
        input_stem: Optional input file stem.

    Returns:
        Path to the generated agent descriptions file, or None on failure.
    """
    print(f"\n{'─' * 40}")
    print("m2: Agent Descriptions File Generation")
    print("─" * 40)

    # ── Locate analysis-manifest.json ──
    analysis_manifest_path = _find_analysis_manifest(
        output_md, images_dir, fmt,
        short_name=short_name, input_stem=input_stem)

    if analysis_manifest_path is None:
        _warn("m2: No analysis-manifest.json found. "
              "Run prepare-image-analysis.py first, then retry "
              "with --agent-descriptions.")
        return None

    # ── Locate image-manifest.json ──
    image_manifest_path = _find_image_manifest(
        output_md, images_dir, fmt,
        short_name=short_name, input_stem=input_stem)

    # ── Parse the image index for substantive entries ──
    substantive_entries = _parse_image_index_file(image_index_path)
    if not substantive_entries:
        print("  No substantive images found in image index. "
              "Nothing to describe.")
        return None

    substantive_pages = {e["page"] for e in substantive_entries}
    print(f"  Substantive images: {len(substantive_entries)} "
          f"across {len(substantive_pages)} page(s)")

    # ── Read analysis manifest for document domain and image details ──
    try:
        analysis_data = json.loads(
            analysis_manifest_path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError) as e:
        _warn(f"m2: Could not parse analysis manifest: {e}")
        return None

    document_domain = analysis_data.get("document_domain", "general")
    document_title = analysis_data.get("document_title", output_md.stem)
    # Normalize embedded newlines/excess whitespace in title
    # (e.g. PPTX titles with line breaks: "HTA Systems &\nClinical Evidence")
    document_title = " ".join(document_title.split())
    analysis_images = analysis_data.get("images", [])
    total_analysis_images = analysis_data.get("total_images", 0)

    # ── Read image manifest for nearby_text and additional context ──
    manifest_images = []
    if image_manifest_path is not None:
        try:
            manifest_data = json.loads(
                image_manifest_path.read_text(encoding="utf-8"))
            manifest_images = manifest_data.get("images", [])
        except (json.JSONDecodeError, OSError) as e:
            _warn(f"m2: Could not parse image manifest: {e}. "
                  "Proceeding without nearby_text data.")

    # Build a lookup of manifest images by page number for nearby_text
    # Key: (page, figure_num) → image entry
    manifest_by_page: dict = {}
    for img in manifest_images:
        page = img.get("page", 0)
        fig = img.get("figure_num", 0)
        manifest_by_page[(page, fig)] = img

    # Build a lookup of analysis images by figure_num
    analysis_by_fig: dict = {}
    for img in analysis_images:
        fig = img.get("figure_num", 0)
        analysis_by_fig[fig] = img

    # ── Resolve the generate-image-notes.md skill path ──
    skill_path = SCRIPTS_DIR / "generate-image-notes.md"
    if not skill_path.exists():
        skill_path_str = str(skill_path)
        _warn(f"m2: generate-image-notes.md not found at {skill_path_str}. "
              "The agent will need to locate it manually.")
    else:
        skill_path_str = str(skill_path)

    # ── Build the output file ──
    # Write alongside the image index in the same directory
    output_dir = image_index_path.parent
    desc_filename = f"{image_index_path.stem.replace('-image-index', '')}" \
                    f"-agent-descriptions.md"
    desc_path = output_dir / desc_filename

    now = datetime.now(timezone.utc)
    timestamp = now.strftime("%Y-%m-%d %H:%M UTC")

    lines = []
    lines.append(f"# Agent Descriptions Prompt: {document_title}")
    lines.append("")
    lines.append(f"Generated: {timestamp}")
    lines.append(f"Pipeline version: v{PIPELINE_VERSION} (m2)")
    lines.append(f"Source image index: {image_index_path}")
    lines.append(f"Analysis manifest: {analysis_manifest_path}")
    if image_manifest_path:
        lines.append(f"Image manifest: {image_manifest_path}")
    lines.append("")
    lines.append("---")
    lines.append("")

    # ── Section 1: Document context ──
    lines.append("## Document Context")
    lines.append("")
    lines.append(f"- **Document title:** {document_title}")
    lines.append(f"- **Document domain:** {document_domain}")
    lines.append(f"- **Format:** {fmt.upper()}")
    lines.append(f"- **MD file:** `{output_md.resolve()}`")
    images_dir_resolved = analysis_data.get(
        "images_dir", str(images_dir.resolve()))
    lines.append(f"- **Images directory:** `{images_dir_resolved}`")
    lines.append(f"- **Total images detected (pre-filter):** "
                 f"{total_analysis_images}")
    # Pre-count actual table rows: each substantive page contributes one row
    # per analysis-manifest image on that page (or 1 fallback row if none).
    _desc_row_count = 0
    for _e in substantive_entries:
        _pg_imgs = [img for img in analysis_images
                    if img.get("page") == _e["page"]]
        _desc_row_count += len(_pg_imgs) if _pg_imgs else 1
    lines.append(f"- **Substantive images to describe:** "
                 f"{_desc_row_count}")
    lines.append("")

    # ── Section 2: File paths for the agent ──
    lines.append("## Files for Agent")
    lines.append("")
    lines.append("The agent must read these files in this order:")
    lines.append("")
    lines.append(f"1. **Full markdown** (document context): "
                 f"`{output_md.resolve()}`")
    if image_manifest_path:
        lines.append(f"2. **Image manifest** (image inventory): "
                     f"`{image_manifest_path.resolve()}`")
    lines.append(f"3. **Analysis manifest** (persona activations): "
                 f"`{analysis_manifest_path.resolve()}`")
    lines.append(f"4. **Skill prompt** (IMAGE NOTE schema): "
                 f"`{skill_path_str}`")
    lines.append("")

    # ── Section 3: Substantive images table ──
    lines.append("## Substantive Images")
    lines.append("")
    lines.append("Each row below is a substantive image that needs "
                 "an IMAGE NOTE description.")
    lines.append("")
    lines.append("| Fig | Page | File Path | Section | "
                 "Nearby Text (first 120 chars) |")
    lines.append("|-----|------|-----------|---------|"
                 "-------------------------------|")

    for entry in substantive_entries:
        page = entry["page"]
        img_count = entry.get("image_count", 1)

        # Find all analysis manifest images on this page
        page_images = [img for img in analysis_images
                       if img.get("page") == page]

        if not page_images:
            # Fallback: create a row from the index entry alone
            ctx_short = entry.get("context", "")[:120]
            ctx_short = ctx_short.replace("|", "\\|")
            lines.append(
                f"| ? | {page} | (see manifest) | "
                f"— | {ctx_short} |")
            continue

        for img in page_images:
            fig_num = img.get("figure_num", "?")
            filename = img.get("filename", "?")
            file_path = img.get("file_path", filename)

            # Section context from analysis manifest
            section_ctx = img.get("section_context", {})
            section_heading = section_ctx.get("heading", "—")
            section_heading = section_heading.replace("|", "\\|")

            # Nearby text: prefer image manifest (has nearby_text),
            # fall back to index entry context
            nearby = ""
            manifest_img = manifest_by_page.get((page, fig_num))
            if manifest_img:
                nearby = manifest_img.get("nearby_text") or ""
            if not nearby:
                nearby = entry.get("context", "")

            nearby_short = nearby[:120].replace("|", "\\|")
            if len(nearby) > 120:
                nearby_short += "..."

            lines.append(
                f"| {fig_num} | {page} | `{file_path}` | "
                f"{section_heading} | {nearby_short} |")

    lines.append("")

    # ── Section 4: Activated personas summary ──
    lines.append("## Persona Activations Summary")
    lines.append("")
    lines.append("Pre-computed persona activations from "
                 "analysis-manifest.json:")
    lines.append("")

    persona_counts: dict = {}
    for img in analysis_images:
        page = img.get("page", 0)
        if page not in substantive_pages:
            continue
        for persona in img.get("activated_personas", []):
            persona_counts[persona] = persona_counts.get(persona, 0) + 1
        for persona in img.get("conditional_personas", {}):
            persona_counts[persona] = (
                persona_counts.get(persona, 0))  # count conditional once

    if persona_counts:
        lines.append("| Persona | Activations |")
        lines.append("|---------|-------------|")
        for persona, count in sorted(persona_counts.items(),
                                     key=lambda x: -x[1]):
            lines.append(f"| {persona} | {count} |")
    else:
        lines.append("No persona activations found in manifest.")
    lines.append("")

    # ── Section 5: Task invocation template ──
    lines.append("## Agent Invocation")
    lines.append("")
    lines.append("Use this Task() call to launch the subagent:")
    lines.append("")
    lines.append("```python")
    lines.append("Task(")
    lines.append('    subagent_type="general-purpose",')
    lines.append(f'    prompt=f"""')
    lines.append(f"    Read the prompt file for full instructions:")
    lines.append(f"    {skill_path_str}")
    lines.append("")
    lines.append(f"    Then generate multi-persona expert IMAGE NOTEs for:")
    lines.append(f"    MD file: {output_md.resolve()}")
    lines.append(f"    Images dir: {images_dir_resolved}")
    if image_manifest_path:
        lines.append(f"    Manifest: {image_manifest_path.resolve()}")
    lines.append(f"    Analysis manifest: "
                 f"{analysis_manifest_path.resolve()}")
    lines.append("")
    lines.append("    PROCEDURE (context-first):")
    lines.append("    1. Read the full markdown file FIRST "
                 "(document context)")
    lines.append("    2. Read image-manifest.json (image inventory)")
    lines.append("    3. Read analysis-manifest.json "
                 "(persona activations pre-computed)")
    lines.append("    4. For each image: Read image via Read tool, "
                 "generate base + persona IMAGE NOTE")
    lines.append("    5. Write all IMAGE NOTEs via Bash+Python")
    lines.append("    6. Update YAML header (image_notes, "
                 "persona_analysis, persona_version, flagged_images)")
    lines.append("    7. Run self-validation checklist")
    lines.append("    8. Report: image count, persona count, "
                 "severity summary only")
    lines.append('    """,')
    lines.append('    description="Generate multi-persona expert '
                 'IMAGE NOTEs"')
    lines.append(")")
    lines.append("```")
    lines.append("")
    lines.append("---")
    lines.append("")
    lines.append(f"*Generated by Pipeline v{PIPELINE_VERSION} (m2) on "
                 f"{timestamp}*")

    # ── Write to disk ──
    try:
        desc_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
        print(f"  GENERATED: {desc_path}")
        print(f"    Substantive images: {len(substantive_entries)}")
        print(f"    Analysis images included: "
              f"{sum(1 for img in analysis_images if img.get('page') in substantive_pages)}")
        print(f"    Document domain: {document_domain}")
        return desc_path
    except Exception as e:
        _warn(f"m2: Could not write agent descriptions file: {e}")
        return None


def generate_testable_index(project_dir: Path) -> Optional[Path]:
    """R20: Generate project-level testable image index.

    Scans all *-image-index.md files in the project directory,
    filters for substantive images, groups by topic, and writes
    TESTABLE-IMAGE-INDEX.md.

    Args:
        project_dir: Root directory of the project.

    Returns:
        Path to the generated TESTABLE-IMAGE-INDEX.md, or None on failure.
    """
    print(f"\n{'=' * 60}")
    print("TESTABLE IMAGE INDEX GENERATION (R20)")
    print(f"Project: {project_dir}")
    print("=" * 60)

    # Find all image index files recursively
    index_files = sorted(project_dir.rglob("*-image-index.md"))

    if not index_files:
        print("  No image index files found in project directory.")
        print("  Run the pipeline on source documents first to generate "
              "per-file indexes.")
        return None

    print(f"  Found {len(index_files)} image index file(s)")

    # Load topic config
    topics, source_patterns = _load_topic_config(project_dir)

    # Parse all index files and collect substantive entries
    all_entries = []
    for idx_file in index_files:
        entries = _parse_image_index_file(idx_file)
        all_entries.extend(entries)

    if not all_entries:
        print("  No substantive images found across all index files.")
        return None

    print(f"  Total substantive image entries: {len(all_entries)}")

    # Classify entries by topic and source category
    # When topics is None, _classify_topic groups by document filename
    use_topic_classification = topics is not None
    topic_groups: dict = {}
    for entry in all_entries:
        topic = _classify_topic(entry["context"],
                                entry["source_file"], topics)
        category = _detect_source_category(entry["source_path"],
                                           source_patterns)
        entry["topic"] = topic
        entry["source_category"] = category
        if use_topic_classification:
            entry["why_testable"] = (
                f"This {topic.lower()} figure from {category.lower()} "
                f"is likely to appear on exams covering "
                f"{topic.lower()} concepts."
            )
        else:
            entry["why_testable"] = (
                f"Substantive figure from {category.lower()} in "
                f"{entry['source_file']}."
            )

        if topic not in topic_groups:
            topic_groups[topic] = []
        topic_groups[topic].append(entry)

    # Sort entries within each topic by source category priority
    for topic_name in topic_groups:
        topic_groups[topic_name].sort(
            key=lambda e: _SOURCE_PRIORITY.get(e["source_category"], 99))

    # Sort topics alphabetically, but put "Uncategorised" last
    sorted_topics = sorted(
        topic_groups.keys(),
        key=lambda t: (1 if t == "Uncategorised" else 0, t))

    # Remove empty topics
    sorted_topics = [t for t in sorted_topics if topic_groups[t]]

    # Determine project name from directory
    project_name = project_dir.name

    # Build the output
    now = datetime.now()
    lines = []
    lines.append(f"# {project_name} - Testable Image Index")
    lines.append(f"Generated: {now.strftime('%Y-%m-%d')}")
    lines.append(f"Source files scanned: {len(index_files)}")
    lines.append(f"Total substantive images: {len(all_entries)}")
    lines.append(f"Topics covered: {len(sorted_topics)}")
    lines.append("")
    lines.append("## How to Use")
    lines.append("1. Pick a topic below")
    lines.append("2. Open the PDF with: `open \"path\"`")
    lines.append("3. Go to the specified page")
    lines.append("4. Study or get tested on the image")
    lines.append("")
    lines.append("**Priority key:**")
    lines.append("- CURRENT SLIDES = highest priority (exam slides)")
    lines.append("- PREVIOUS SLIDES = high priority (overlapping content)")
    lines.append("- LITERATURE = medium priority (reference figures)")
    lines.append("- WORKING GROUP = medium priority (applied examples)")
    lines.append("")
    lines.append("---")

    for topic_name in sorted_topics:
        entries = topic_groups[topic_name]
        lines.append("")
        lines.append(f"## {topic_name}")
        lines.append("")

        for entry in entries:
            ctx_short = entry["context"][:80]
            if len(entry["context"]) > 80:
                ctx_short += "..."
            lines.append(f"### {ctx_short}")
            lines.append(f"- **Source:** {entry['source_category']}")
            lines.append(f"- **File:** `{entry['source_file']}`")
            lines.append(f"- **Full path:** `{entry['source_path']}`")
            lines.append(f"- **Page:** {entry['page']}")
            lines.append(f"- **Why testable:** {entry['why_testable']}")
            lines.append("")

    lines.append("---")
    lines.append("")
    lines.append(f"*Generated by Pipeline v{PIPELINE_VERSION} on "
                 f"{now.strftime('%Y-%m-%d %H:%M')}*")

    # Write output
    output_dir = project_dir / "study-outputs" / "image-inventories"
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / "TESTABLE-IMAGE-INDEX.md"

    try:
        output_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
        print(f"\n  GENERATED: {output_path}")
        print(f"    Topics: {len(sorted_topics)}")
        print(f"    Images: {len(all_entries)}")
        print(f"    Files scanned: {len(index_files)}")
        return output_path
    except Exception as e:
        _warn(f"Could not write testable index: {e}")
        return None


# ═══════════════════════════════════════════════════════════════════════════
# POST-PROCESSING: Regex-based text cleanup applied to extracted markdown
# ═══════════════════════════════════════════════════════════════════════════

def _post_process_markdown(md_text: str) -> str:
    """Apply regex-based text cleanup passes to extracted markdown.

    These fixes target known OCR artifacts produced by MinerU on specific
    document types (slide decks with icons, health-economics PDFs, LaTeX
    source documents).  Each pass is independent and idempotent.

    Passes applied:
      F11b – Bold ** marker stripping from heading lines
      F8   – CJK artifact filter (isolated + clustered CJK in Latin context,
              Erasmus logo footer artifacts)
      F9   – Garbled OCR block detection with warning markers
      F12  – Euro symbol restoration ("AC" → "€" before digits)
      F15  – Table caption "Table" prefix restoration
      m11  – LaTeX accent notation → Unicode
    """

    # ── F11b: Bold ** marker stripping from headings ────────────────────
    # pymupdf4llm wraps heading text in ** bold markers, producing lines
    # like "# **When**" or "## **Where** **to search**".  Strip all **
    # from heading lines so downstream processing sees clean headings.
    # Applied FIRST so headings are clean for subsequent passes.
    def _strip_heading_bold(line: str) -> str:
        if re.match(r'^#{1,6}\s', line):
            return line.replace('**', '')
        return line

    md_text = '\n'.join(
        _strip_heading_bold(ln) for ln in md_text.split('\n')
    )

    # ── F8: CJK artifact filter ──────────────────────────────────────────
    # MinerU OCR misreads decorative slide icons/arrows as CJK characters.
    # Two passes: isolated single char, then clusters of 2-4 chars.
    # Only removes CJK that is surrounded by Latin text / whitespace —
    # leaves intact any genuine CJK content (e.g. Chinese-language PDFs).

    # F8a: single isolated CJK character between Latin chars/spaces
    md_text = re.sub(
        r'(?<=[a-zA-Z\s])[\u4e00-\u9fff](?=[a-zA-Z\s])',
        '',
        md_text,
    )
    # F8b: cluster of 2-4 consecutive CJK characters between Latin chars/spaces
    md_text = re.sub(
        r'(?<=[a-zA-Z\s])[\u4e00-\u9fff]{2,4}(?=[a-zA-Z\s])',
        '',
        md_text,
    )
    # F8c: Erasmus logo footer artifacts ("Erafmy", "Ezamy", variants)
    md_text = re.sub(r'E[rz]a[fm]y', '', md_text, flags=re.IGNORECASE)

    # ── F9: Garbled OCR block detection with warning markers ─────────────
    # MinerU sometimes OCRs overlapping text boxes or chart sub-panels,
    # producing unreadable output like "expresserdinthee rate at thef".
    # Detect paragraphs with a high ratio of garbled-word indicators and
    # prepend a visible warning so reviewers can locate them easily.

    def _is_garbled_word(w: str) -> bool:
        """True if a word looks like garbled OCR output."""
        # Mixed alphanumeric in unusual pattern (letters+digits+letters)
        if re.search(r'[a-zA-Z]{2,}\d+[a-zA-Z]', w):
            return True
        # Very long word with no vowels (consonant soup)
        if len(w) > 8 and not re.search(r'[aeiouAEIOU]', w):
            return True
        return False

    def _process_paragraph(para: str) -> str:
        """Return para unchanged, or with OCR warning prepended."""
        words = para.split()
        if len(words) < 5:
            return para
        garbled_count = sum(1 for w in words if _is_garbled_word(w))
        if garbled_count / len(words) > 0.3:
            return (
                "[OCR garbled — refer to source document]\n" + para
            )
        return para

    # Split on double newlines to process paragraph by paragraph,
    # then rejoin preserving the original blank-line structure.
    paragraphs = md_text.split('\n\n')
    paragraphs = [_process_paragraph(p) for p in paragraphs]
    md_text = '\n\n'.join(paragraphs)

    # ── F12: Euro symbol restoration ─────────────────────────────────────
    # MinerU OCR renders "€" as "AC" in some health-economics PDFs.
    # Pattern: "AC" immediately followed by a digit (with optional space).
    # False-positive risk is negligible — "AC 25" is not a real English
    # phrase, and the pattern requires the digit after to anchor it.
    md_text = re.sub(r'\bAC\s?(\d)', r'€\1', md_text)

    # ── F15: Table caption "Table" prefix restoration ─────────────────────
    # MinerU strips the word "Table" from some captions, leaving e.g.
    # "42: Analyses performed" instead of "Table 42: Analyses performed".
    # Pattern: line starts with 1-3 digits + colon + space + uppercase.
    # Specific enough to avoid false positives (numbered lists already
    # use a period or parenthesis, not a colon, after the number).
    md_text = re.sub(
        r'^(\d{1,3}): ([A-Z])',
        r'Table \1: \2',
        md_text,
        flags=re.MULTILINE,
    )

    # ── m11: LaTeX accent notation → Unicode ─────────────────────────────
    # MinerU sometimes preserves raw LaTeX notation from source documents
    # (e.g. "Ren\'ee" instead of "Renée").  Replace common sequences.
    LATEX_ACCENTS = {
        r"\'e": "é", r"\'a": "á", r"\'i": "í",
        r"\'o": "ó", r"\'u": "ú", r"\'E": "É",
        r"\`e": "è", r"\`a": "à",
        r'\"o': "ö", r'\"u': "ü", r'\"a': "ä",
        r"\~n": "ñ",
    }
    for latex_seq, uni_char in LATEX_ACCENTS.items():
        md_text = md_text.replace(latex_seq, uni_char)
    # Handle the apostrophe-only variant (without backslash):
    # e.g. "'ee" → "ée" and "'e" at word boundary → "é"
    md_text = re.sub(r"(\w)'ee\b", r"\1ée", md_text)
    md_text = re.sub(r"(\w)'e\b", r"\1é", md_text)

    # ── RC8: MinerU backslash escaping cleanup ────────────────────────────
    # MinerU wraps ~ and * in unnecessary backslash escapes (\~ and \*).
    # Must run AFTER m11 (LaTeX accents) to avoid stripping legitimate
    # LaTeX sequences like \~n → ñ (already handled above).
    md_text = re.sub(r'\\([~*])', r'\1', md_text)

    # ── RC9: MinerU Japanese period (U+3002) → ASCII period ──────────────
    # MinerU OCR sometimes misreads decimal points as Japanese full-stops
    # (。).  Context-aware: only replace when between digits (with optional
    # whitespace), preserving any genuine CJK punctuation.
    md_text = re.sub(r'(?<=\d)\u3002(?=\s*\d)', '.', md_text)

    return md_text


# ═══════════════════════════════════════════════════════════════════════════
# FIX 3.12: Image-Markdown Sync
# ═══════════════════════════════════════════════════════════════════════════

def sync_images_to_md(md_path: Path, manifest_path: Path) -> bool:
    """Replace 'No images extracted' placeholder with image index.

    Fix 3.12: After image extraction, the manifest may contain images but
    the markdown text may still say 'No images extracted.'  This function
    scans the MD for that placeholder and replaces it with a proper image
    index table built from the manifest.

    Returns True if the placeholder was found and replaced, False otherwise.
    """
    if not md_path.exists() or not manifest_path.exists():
        return False

    md_content = md_path.read_text(encoding='utf-8')
    if 'No images extracted' not in md_content:
        return False  # Nothing to fix

    manifest = json.loads(manifest_path.read_text(encoding='utf-8'))
    images = manifest.get("images", [])
    if not images:
        return False

    # Build image index table (no heading — the MD already has ## Image Index)
    # Filter to substantive images only (backward compat: keep all if fields missing)
    substantive = []
    excluded_count = 0
    for img in images:
        has_filter_fields = any(
            k in img for k in ("is_substantive", "is_blank", "is_duplicate")
        )
        if has_filter_fields:
            if img.get("is_blank", False) or img.get("is_duplicate", False):
                excluded_count += 1
                continue
            if "is_substantive" in img and not img.get("is_substantive", True):
                excluded_count += 1
                continue
        substantive.append(img)

    index_lines = ["| Figure | File | Page | Size |",
                   "|--------|------|------|------|"]
    for img in substantive:
        fn = img.get("filename", "")
        pg = img.get("page", "?")
        w = img.get("width", "?")
        h = img.get("height", "?")
        fig = img.get("figure_num", "?")
        index_lines.append(f"| {fig} | {fn} | {pg} | {w}x{h} |")

    if excluded_count > 0:
        index_lines.append(
            f"\n*Showing {len(substantive)} substantive images. "
            f"{excluded_count} decorative/blank/duplicate images excluded.*"
        )

    index_text = "\n".join(index_lines)
    md_content = md_content.replace(
        "No images extracted.", index_text)
    md_path.write_text(md_content, encoding='utf-8')
    return True


# ═══════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(
        description="Pipeline orchestrator with extractor selection router",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python3 run-pipeline.py paper.pdf
  python3 run-pipeline.py report.docx -o reference-papers/report.md
  python3 run-pipeline.py slides.pptx --no-images
  python3 run-pipeline.py scanned.pdf --force-extractor tesseract
        """
    )

    parser.add_argument("input_file", type=Path, nargs="?", default=None,
                        help="Input file (PDF/DOCX/PPTX/TXT). "
                             "Optional when --generate-testable-index is used "
                             "standalone.")
    parser.add_argument("--output", "-o", type=Path, default=None,
                        help="Output markdown path (default: <input>.md)")
    parser.add_argument("--images-dir", "-i", type=Path, default=None,
                        help="Images directory "
                             "(default: images/<short-name>/)")
    parser.add_argument("--no-images", action="store_true",
                        help="Skip image extraction")
    parser.add_argument("--short-name", "-s", type=str, default=None,
                        help="Short name for filenames "
                             "(default: from input name)")
    parser.add_argument("--skip-numbers", action="store_true",
                        help="Skip number extraction (for non-PDF)")
    parser.add_argument("--force-extractor", type=str, default=None,
                        choices=["pymupdf4llm", "tesseract", "mineru",
                                 "zerox", "docling", "marker"],
                        help="Force a specific extractor "
                             "(bypass auto-detection)")
    parser.add_argument("--skip-cross-validation", action="store_true",
                        help="Skip pdfplumber cross-validation")

    # ── v3.1 flags ────────────────────────────────────────────────────────
    # R8: --target-dir activates auto-organization (R1-R7).
    # When omitted, pipeline behaves identically to v3.0 (R16).
    parser.add_argument("--target-dir", "-t", type=Path, default=None,
                        metavar="PATH",
                        help="Target directory for organized output. "
                             "Created if it does not exist. "
                             "When provided: activates auto-organization "
                             "(moves source to _originals/, places .md and "
                             "images, writes visual report). "
                             "When omitted: v3.0 behavior, no organization.")
    parser.add_argument("--force", action="store_true",
                        help="Bypass duplicate detection and re-convert "
                             "even if a registry entry exists for this file.")
    parser.add_argument("--organize-only", action="store_true",
                        help="Skip conversion steps (0-6) and run only the "
                             "organization phase (Steps 7-13). Useful when "
                             "conversion succeeded but organization failed.")
    parser.add_argument("--dry-run", action="store_true",
                        help="Print what WOULD happen but perform no file "
                             "operations. Shows the visual report with "
                             "(DRY RUN) markers on all moved/placed/deleted "
                             "lines. Requires --target-dir.")
    parser.add_argument("--generate-testable-index", type=Path, default=None,
                        metavar="PROJECT_DIR",
                        help="R20: Generate a project-level testable image "
                             "index by aggregating all *-image-index.md files "
                             "in the given project directory. Writes to "
                             "[project-dir]/study-outputs/image-inventories/"
                             "TESTABLE-IMAGE-INDEX.md. Can be combined with "
                             "a normal conversion or run standalone.")

    # ── m2: --agent-descriptions flag ─────────────────────────────────────
    # When provided alongside a conversion or with --generate-testable-index,
    # generates an agent prompt file containing all information needed to
    # describe substantive images using the generate-image-notes.md skill.
    # The output file is written alongside the image index / analysis manifest.
    parser.add_argument("--agent-descriptions", action="store_true",
                        default=False,
                        help="m2: After conversion, generate an agent "
                             "descriptions prompt file that provides "
                             "everything a Claude subagent needs to "
                             "describe substantive images (paths, page "
                             "numbers, nearby text, document domain). "
                             "Requires an image index and analysis "
                             "manifest to exist. Can be combined with "
                             "conversion or run standalone with an "
                             "input file.")

    # ── Issue reporting flags ──────────────────────────────────────────────
    parser.add_argument('--report-issue', action='store_true',
                        help='Interactively write a pipeline issue report '
                             'to the reports folder')
    parser.add_argument('--health-check', action='store_true',
                        help='Read all pipeline reports and show unresolved '
                             'issues')

    args = parser.parse_args()

    # ── Issue reporting early exits ────────────────────────────────────────
    if args.health_check:
        run_health_check()
        sys.exit(0)

    if args.report_issue:
        interactive_report_issue()
        sys.exit(0)

    # ── R20: Standalone --generate-testable-index mode ────────────────────
    # When --generate-testable-index is provided WITHOUT an input file,
    # run project-level aggregation only and exit.
    if args.generate_testable_index is not None and args.input_file is None:
        # m2: --agent-descriptions requires an input file (per-file feature)
        if args.agent_descriptions:
            _warn("--agent-descriptions requires an input file. "
                  "It is a per-file feature and cannot run in "
                  "standalone --generate-testable-index mode. "
                  "Ignoring --agent-descriptions.")
        project_dir = args.generate_testable_index.resolve()
        if not project_dir.exists():
            print(f"ERROR: Project directory not found: {project_dir}",
                  file=sys.stderr)
            sys.exit(1)
        result = generate_testable_index(project_dir)
        if result:
            print(f"\nTestable index written to: {result}")
            sys.exit(0)
        else:
            print("\nNo testable index generated.", file=sys.stderr)
            sys.exit(1)

    # ── Validate input_file is provided for conversion modes ──────────────
    if args.input_file is None:
        print("ERROR: input_file is required unless using "
              "--generate-testable-index standalone.", file=sys.stderr)
        sys.exit(1)

    if not args.input_file.exists():
        if args.organize_only and args.target_dir:
            # Source may have been moved to _originals/ by a previous partial run
            _originals_candidate = (
                args.target_dir.resolve() / ORIGINALS_SUBDIR
                / args.input_file.name
            )
            print(f"ERROR: Input file not found: {args.input_file}")
            print("  With --organize-only, the source may have been moved")
            print("  by a previous partial run. If so, pass the current path:")
            print(f"    {_originals_candidate}")
        else:
            print(f"ERROR: Input file not found: {args.input_file}")
        sys.exit(1)

    # ── Path-space safety: resolve and validate input_file ────────────────
    # When the input path contains spaces and --target-dir is used, shell
    # quoting issues can cause argparse to receive a truncated path (e.g.
    # PosixPath('.')) which has an empty name and suffix, causing
    # ValueError on .with_suffix().  Resolve early and validate.
    args.input_file = args.input_file.resolve()
    if not args.input_file.name or args.input_file.name == ".":
        print("ERROR: Input file path resolved to an invalid path "
              f"({args.input_file}).\n"
              "  This usually means the file path contains spaces and "
              "was not properly quoted.\n"
              '  Fix: wrap the path in quotes, e.g.:\n'
              '    python3 run-pipeline.py "path with spaces/file.pptx" '
              '--target-dir /tmp/out',
              file=sys.stderr)
        sys.exit(1)

    # Determine format and paths
    fmt = args.input_file.suffix.lower().lstrip(".")
    output_md = args.output or args.input_file.with_suffix(".md")
    short_name = (args.short_name
                  or args.input_file.stem.lower().replace(" ", "-"))
    images_dir = (args.images_dir
                  or (output_md.parent / "images" / short_name))

    # ── Compute source hash ONCE, early ─────────────────────────────────
    # SHA-256 is needed for registry updates, duplicate detection, and
    # organization tracking.  Compute it here while args.input_file is
    # guaranteed to exist (before Step 8 moves it to _originals/).
    source_hash = _compute_sha256(args.input_file)

    # ── R13: Excel/XLSX graceful skip ─────────────────────────────────────
    # XLSX files are NEVER converted (lossy: formulas become strings).
    # With --target-dir: move to _originals/ unchanged, log INFO, exit 0.
    # Without --target-dir: print skip message, exit 0.
    # No registry entry written for Excel files.
    if fmt in ("xlsx", "xlsm"):
        print(f"INFO: {fmt.upper()} files are not converted by the pipeline.")
        print("  XLSX/XLSM→MD conversion is lossy (formulas become strings).")
        print("  Use openpyxl to query Excel files directly.")
        if args.target_dir:
            target_dir = args.target_dir.resolve()
            if not args.dry_run:
                target_dir.mkdir(parents=True, exist_ok=True)
            originals_dir = target_dir / ORIGINALS_SUBDIR
            original_dest = originals_dir / args.input_file.name

            if args.dry_run:
                print(f"  (DRY RUN) Would move {args.input_file.name} "
                      f"→ {original_dest}")
            elif original_dest.exists():
                print(f"  SKIP: {args.input_file.name} already in "
                      f"{originals_dir}")
            else:
                try:
                    atomic_move(args.input_file, original_dest)
                    print(f"  MOVED: {args.input_file.name} → {originals_dir}")
                except Exception as e:
                    print(f"  ✗ MOVE FAILED: {e}", file=sys.stderr)
                    sys.exit(3)

            # R6: Log INFO entry for Excel skip
            if not args.dry_run:
                append_issue_log(
                    target_dir=target_dir,
                    source_file=args.input_file,
                    output_md=None,
                    extractor="none",
                    issue_type="INFO",
                    severity="INFO",
                    details=(f"Excel file moved to {ORIGINALS_SUBDIR} without "
                             "conversion (Excel files are not converted)."),
                    action_taken="Moved to _originals/ without conversion.",
                )
        sys.exit(0)

    # ── v3.1 flag validation ──────────────────────────────────────────────
    if args.dry_run and not args.target_dir:
        print("ERROR: --dry-run requires --target-dir.", file=sys.stderr)
        sys.exit(1)
    if args.organize_only and not args.target_dir:
        print("ERROR: --organize-only requires --target-dir.",
              file=sys.stderr)
        sys.exit(1)

    # Checkpoint for failure tracking
    checkpoint_path = output_md.parent / ".pipeline-checkpoint.json"
    checkpoint = {
        "source_file": str(args.input_file.resolve()),
        "output_file": str(output_md.resolve()),
        "format": fmt,
        "current_state": "started",
        "started_at": datetime.now().isoformat(),
    }

    # ── --organize-only: skip conversion, run organization only ──────────
    # R8: When set, skip Steps 0-6 entirely. Assumes conversion already ran.
    # Initialize variables that the organization phase needs, then jump
    # past the conversion block.
    if args.organize_only:
        print("Pipeline Orchestrator v3.1 (--organize-only mode)")
        print(f"Input:  {args.input_file}")
        print(f"Format: {fmt.upper()}")
        print(f"Output: {output_md}")
        print("Skipping Steps 0-6 (conversion). Running organization only.")

        # Sync output_md for office formats (same logic as BUG-4 FIX)
        if fmt in ("docx", "pptx", "txt"):
            actual_office_md = output_md.parent / f"{args.input_file.stem}.md"
            if actual_office_md.exists() and actual_office_md != output_md:
                output_md = actual_office_md

        extractor_config = None
        cross_val_flags = []
        _image_index_meta = None  # R19: not generated in organize-only mode
        input_stem_org = args.input_file.stem
        # Discover manifest for the organization phase
        _manifest_candidates = [
            images_dir / "image-manifest.json",
            output_md.parent / f"{input_stem_org}_manifest.json",
            output_md.parent / f"{short_name}_manifest.json",
        ]
        manifest_path = next(
            (p for p in _manifest_candidates if p.exists()),
            _manifest_candidates[0],
        )
        # R19: Try to discover existing image index for organize-only mode
        _idx_candidate = output_md.parent / f"{output_md.stem}-image-index.md"
        if _idx_candidate.exists():
            # Parse basic metadata from existing image index file
            try:
                _idx_content = _idx_candidate.read_text(encoding="utf-8")
                _idx_meta = {
                    "image_index_path": str(_idx_candidate),
                    "image_index_generated_at": datetime.now().isoformat(),
                    "total_pages": 0,
                    "pages_with_images": 0,
                    "total_images_detected": 0,
                    "substantive_images": 0,
                    "has_testable_images": False,
                }
                for _line in _idx_content.splitlines():
                    if _line.startswith("Total pages:"):
                        _idx_meta["total_pages"] = int(
                            _line.split(":")[1].strip())
                    elif _line.startswith("Pages with images:"):
                        _idx_meta["pages_with_images"] = int(
                            _line.split(":")[1].strip())
                    elif _line.startswith("Total images detected:"):
                        _idx_meta["total_images_detected"] = int(
                            _line.split(":")[1].strip())
                    elif _line.startswith("Estimated substantive images:"):
                        _val = _line.split(":")[1].strip()
                        _idx_meta["substantive_images"] = int(
                            _val.split()[0])
                        _idx_meta["has_testable_images"] = int(
                            _val.split()[0]) > 0
                _image_index_meta = _idx_meta
            except Exception:
                pass  # Non-fatal: organize-only can proceed without it

    # ── End --organize-only setup. Fall through to organization phase. ──

    if not args.organize_only:
        print("Pipeline Orchestrator v3.0")
        print(f"Input:  {args.input_file}")
        print(f"Format: {fmt.upper()}")
        print(f"Output: {output_md}")
        print(f"Images: {images_dir}")

    # ── Steps 0-6: Conversion (skipped when --organize-only) ────────
    if not args.organize_only:
        # ── Step 0: Extractor Selection Router (PDF only) ──
        extractor_config = None
        if fmt == "pdf":
            print(f"\n{'─' * 40}")
            print("Step 0: Extractor Selection Router")
            print('─' * 40)
            try:
                extractor_config = select_extractor(
                    args.input_file,
                    force_extractor=args.force_extractor,
                )
                checkpoint["extractor"] = extractor_config.extractor
                checkpoint["is_scanned"] = extractor_config.is_scanned
                checkpoint["avg_chars_per_page"] = (
                    extractor_config.avg_chars_per_page
                )
                print(f"  Selected: {extractor_config.extractor}")
                if extractor_config.is_scanned:
                    print("  Document type: SCANNED")
                else:
                    print("  Document type: DIGITAL")
            except Exception as e:
                _warn(f"Scan detection failed: {e}. "
                      "Assuming digital, using pymupdf4llm.")
                check_for_known_failures(str(e), context="Step 0: Extractor Selection")
                extractor_config = ExtractorConfig(
                    extractor="pymupdf4llm",
                    script=str(SCRIPTS_DIR / "convert-paper.py"),
                    extra_args=["--extractor", "pymupdf4llm"],
                    is_scanned=False,
                )

        # ── Re-run idempotency: clear stale artifacts ──────────────
        # On re-run, the manifest from a previous Step 6c (with vector
        # render entries) persists, but Step 1 resets the embedded
        # ## Image Index section in the .md to "No images extracted".
        # This causes a manifest/index count mismatch in Step 2 (QC).
        # Fix: delete stale manifest and image index before Step 1 so
        # every run starts clean.  Step 1 recreates the manifest for
        # raster images; Step 6c recreates it for vector renders.
        _stale_manifest = images_dir / "image-manifest.json"
        if _stale_manifest.exists():
            try:
                _stale_manifest.unlink()
                print(f"  Cleared stale manifest: {_stale_manifest.name}")
            except OSError:
                pass  # non-fatal; QC may still warn
        _stale_index = output_md.parent / f"{output_md.stem}-image-index.md"
        if _stale_index.exists():
            try:
                _stale_index.unlink()
                print(f"  Cleared stale image index: {_stale_index.name}")
            except OSError:
                pass

        # ── Step 1: Convert ──
        if fmt == "pdf" and extractor_config:
            # Use centralized command builder
            cmd = _build_cmd_for_extractor(
                extractor_config.extractor, args.input_file,
                output_md, images_dir, short_name, args.no_images,
            )
        else:
            # Non-PDF: route office formats to convert-office.py
            cmd = [
                sys.executable,
                str(Path.home() / ".claude/scripts/convert-office.py"),
                str(args.input_file),
                "--output-dir", str(output_md.parent),
            ]

        # Marker wrapper has internal 600s timeout per attempt (1,200s
        # total with CPU retry).  Outer timeout of 1,500s provides
        # headroom for postprocessing (NFC, ligatures, run-togethers).
        _step1_timeout = (1500 if (extractor_config
                                   and extractor_config.extractor == "marker")
                          else None)
        exit_code = run_command(cmd, "Step 1: Text + Image Extraction",
                               timeout=_step1_timeout)

        # ── BUG-4 FIX: sync output_md for office formats ──
        # convert-office.py always names its output {input_stem}.md in
        # the output directory.  When the user passes --output with a
        # different basename (e.g. --output /tmp/out/costing-pptx.md),
        # output_md points to a file that was never written, causing
        # Step 2 and the manifest search to fail.
        # Fix: after a successful office extraction, update output_md and
        # checkpoint to the path that convert-office.py actually wrote.
        if exit_code == 0 and fmt in ("docx", "pptx", "txt"):
            actual_office_md = output_md.parent / f"{args.input_file.stem}.md"
            if actual_office_md.exists() and actual_office_md != output_md:
                print(f"\nINFO: output_md updated to match convert-office.py "
                      f"output:\n  expected: {output_md}\n"
                      f"  actual:   {actual_office_md}")
                output_md = actual_office_md
                checkpoint["output_file"] = str(output_md.resolve())

        # ── BUG-4b FIX: sync output_md for marker extractor ──
        # convert-paper-marker.py always writes {input_stem}.md to the
        # output directory.  When --output specifies a different basename,
        # output_md points to a file that marker never writes.  Same bug
        # class as BUG-4 for office formats above.
        if (exit_code == 0 and fmt == "pdf"
                and extractor_config
                and extractor_config.extractor == "marker"):
            actual_marker_md = output_md.parent / f"{args.input_file.stem}.md"
            if actual_marker_md.exists() and actual_marker_md != output_md:
                print(f"\nINFO: output_md updated to match marker output:\n"
                      f"  expected: {output_md}\n"
                      f"  actual:   {actual_marker_md}")
                output_md = actual_marker_md
                checkpoint["output_file"] = str(output_md.resolve())

        # ── R19: Discover image index written by convert-office.py ──
        # convert-office.py generates {stem}-image-index.md in the output
        # directory.  Populate _image_index_meta so the organization phase
        # (Step 9, lines ~3008-3032) can move it to --target-dir and the
        # registry call (Bug 2 fix) can include R21 fields.
        _office_image_index_meta = None  # Initialize before conditional
        if exit_code == 0 and fmt in ("docx", "pptx"):
            _idx_candidate = (
                output_md.parent / f"{output_md.stem}-image-index.md"
            )
            if _idx_candidate.exists():
                try:
                    _idx_content = _idx_candidate.read_text(
                        encoding="utf-8")
                    _office_idx_meta = {
                        "image_index_path": str(_idx_candidate),
                        "image_index_generated_at": (
                            datetime.now(timezone.utc).isoformat()),
                        "total_pages": 0,
                        "pages_with_images": 0,
                        "total_images_detected": 0,
                        "substantive_images": 0,
                        "has_testable_images": False,
                    }
                    for _line in _idx_content.splitlines():
                        if _line.startswith("Total pages:"):
                            _office_idx_meta["total_pages"] = int(
                                _line.split(":")[1].strip())
                        elif _line.startswith("Pages with images:"):
                            _office_idx_meta["pages_with_images"] = int(
                                _line.split(":")[1].strip())
                        elif _line.startswith(
                                "Total images detected:"):
                            _office_idx_meta[
                                "total_images_detected"] = int(
                                _line.split(":")[1].strip())
                        elif _line.startswith(
                                "Estimated substantive images:"):
                            _val = _line.split(":")[1].strip()
                            _office_idx_meta[
                                "substantive_images"] = int(
                                _val.split()[0])
                            _office_idx_meta[
                                "has_testable_images"] = (
                                int(_val.split()[0]) > 0)
                    # Store for downstream use (registry + move)
                    _office_image_index_meta = _office_idx_meta
                    print(f"\n  R19: Discovered image index from "
                          f"convert-office.py: {_idx_candidate.name}")
                    print(f"    Pages with images: "
                          f"{_office_idx_meta['pages_with_images']}")
                    print(f"    Substantive: "
                          f"{_office_idx_meta['substantive_images']}")

                    # ── FIX-6: Reconcile image count with manifest ──
                    # The DOCX/PPTX image index counts drawing XML
                    # elements, which can exceed actual extracted images
                    # (e.g. non-extractable shapes, duplicated refs).
                    # Cross-check against manifest and correct if needed.
                    _input_stem_fix6 = args.input_file.stem
                    _manifest_count_fix6 = _read_image_count_from_manifest(
                        output_md.parent, _input_stem_fix6)
                    _actual_images = _manifest_count_fix6[1]  # unique
                    _idx_detected = _office_idx_meta[
                        "total_images_detected"]
                    _fix6_content = None
                    if (_actual_images > 0
                            and _idx_detected != _actual_images):
                        print(f"  FIX-6: Index says "
                              f"{_idx_detected} images but manifest "
                              f"has {_actual_images}. Correcting.")
                        _office_idx_meta[
                            "total_images_detected"] = _actual_images
                        _fix6_content = _idx_candidate.read_text(
                            encoding="utf-8")
                        _fix6_content = re.sub(
                            r"^Total images detected: \d+",
                            f"Total images detected: {_actual_images}",
                            _fix6_content,
                            flags=re.MULTILINE,
                        )
                    # ── F7: Bounds-cap substantive_images independently ──
                    # substantive_images can exceed the unique manifest count
                    # even when total_images_detected already matches (e.g.
                    # DOCX images at section boundaries counted twice in the
                    # XML scan).  Apply the cap regardless of whether the
                    # total count needed correction.
                    if _actual_images > 0:
                        _cur_sub = _office_idx_meta.get(
                            "substantive_images", 0)
                        _capped_sub = min(_cur_sub, _actual_images)
                        if _capped_sub != _cur_sub:
                            print(f"  F7: Substantive count capped "
                                  f"{_cur_sub} → {_capped_sub} "
                                  f"(unique manifest images: "
                                  f"{_actual_images})")
                            _office_idx_meta[
                                "substantive_images"] = _capped_sub
                            _office_idx_meta[
                                "has_testable_images"] = _capped_sub > 0
                            if _fix6_content is None:
                                _fix6_content = _idx_candidate.read_text(
                                    encoding="utf-8")
                            _fix6_content = re.sub(
                                r"^(Estimated substantive images:) \d+",
                                rf"\1 {_capped_sub}",
                                _fix6_content,
                                flags=re.MULTILINE,
                            )
                    # Write corrected content to disk if anything changed
                    if _fix6_content is not None:
                        _idx_candidate.write_text(
                            _fix6_content, encoding="utf-8")
                except Exception as _e:
                    _warn(f"Could not parse office image index: {_e}")
                    _office_image_index_meta = None
            else:
                _office_image_index_meta = None

        # ── Runtime fallback chain ──
        # If the selected extractor fails at runtime, try the next
        # extractor in the appropriate chain before giving up.
        #
        # Scanned chain: tesseract -> mineru -> zerox
        # Digital chain: marker -> docling -> pymupdf4llm -> mineru -> tesseract
        #
        # BUG FIX: the original code only ran fallbacks when is_scanned
        # was True. Digital PDFs that cause pymupdf4llm to crash (e.g.
        # ValueError in pymupdf/table.py line 1534 when a table has an
        # empty bounding-box list) were left with no fallback and the
        # pipeline immediately reported "all extractors exhausted".
        if exit_code != 0 and extractor_config:
            current = extractor_config.extractor
            while exit_code != 0:
                if extractor_config.is_scanned:
                    fallback = _next_scanned_fallback(current)
                    if fallback is None:
                        # Forced extractor not in scanned chain — try digital chain
                        fallback = _next_digital_fallback(current)
                else:
                    fallback = _next_digital_fallback(current)
                if fallback is None:
                    break
                _warn(f"{current} failed at runtime. "
                      f"Trying fallback: {fallback}")
                # Determine script path for fallback extractor
                if fallback == "marker":
                    _fb_script = _MARKER_WRAPPER
                elif fallback == "mineru":
                    _fb_script = str(SCRIPTS_DIR / "convert-mineru.py")
                elif fallback == "zerox":
                    _fb_script = str(
                        SCRIPTS_DIR / f"convert-{fallback}.py")
                else:
                    _fb_script = str(SCRIPTS_DIR / "convert-paper.py")
                extractor_config = ExtractorConfig(
                    extractor=fallback,
                    script=_fb_script,
                    extra_args=(["--extractor", fallback]
                                if fallback in ("pymupdf4llm", "tesseract",
                                                "docling")
                                else []),
                    is_scanned=extractor_config.is_scanned,
                    avg_chars_per_page=(
                        extractor_config.avg_chars_per_page),
                    page_count=extractor_config.page_count,
                )
                checkpoint["extractor"] = fallback
                checkpoint["fallback_chain"] = checkpoint.get(
                    "fallback_chain", []) + [current]
                cmd = _build_cmd_for_extractor(
                    fallback, args.input_file,
                    output_md, images_dir, short_name, args.no_images,
                )
                _fb_timeout = 1500 if fallback == "marker" else None
                exit_code = run_command(
                    cmd,
                    f"Step 1 (fallback): {fallback}",
                    timeout=_fb_timeout,
                )
                current = fallback

        if exit_code != 0:
            _fail("Step 1 (convert) failed - all extractors exhausted",
                  checkpoint, checkpoint_path)

        checkpoint["current_state"] = "extraction_complete"

        # ── Step 1b: Cross-Validation (PDF + pymupdf4llm only) ──
        cross_val_flags = []
        _is_pymupdf_cv = (fmt == "pdf"
                          and extractor_config
                          and extractor_config.extractor == "pymupdf4llm"
                          and not args.skip_cross_validation)
        if (fmt == "pdf" and extractor_config
                and extractor_config.extractor != "pymupdf4llm"
                and not args.skip_cross_validation):
            print(f"\n  Step 1b: Cross-validation skipped "
                  f"(not applicable for {extractor_config.extractor} "
                  f"extractor)")

            # ── Quality gate: word-count ratio vs fitz ──
            # Full cross-validation only exists for pymupdf4llm (uses
            # pdfplumber as reference). For other extractors, compare
            # word count against fitz raw text as a lightweight check.
            if extractor_config.extractor in ("docling", "marker") and output_md.exists():
                _gate_name = extractor_config.extractor.capitalize()
                _gate_critical = _extractor_quality_gate(
                    _gate_name, output_md, args.input_file)
                if _gate_critical:
                    # Critically empty output (<10%): re-enter fallback
                    # chain from the current extractor position.
                    exit_code = 1
                    print(f"  Quality gate CRITICAL: forcing fallback")
                    current = extractor_config.extractor
                    while exit_code != 0:
                        fallback = _next_digital_fallback(current)
                        if fallback is None:
                            break
                        _warn(f"{current} quality gate failed. "
                              f"Trying fallback: {fallback}")
                        if fallback == "marker":
                            _fb_script = _MARKER_WRAPPER
                        elif fallback == "mineru":
                            _fb_script = str(
                                SCRIPTS_DIR / "convert-mineru.py")
                        elif fallback == "zerox":
                            _fb_script = str(
                                SCRIPTS_DIR / f"convert-{fallback}.py")
                        else:
                            _fb_script = str(
                                SCRIPTS_DIR / "convert-paper.py")
                        extractor_config = ExtractorConfig(
                            extractor=fallback,
                            script=_fb_script,
                            extra_args=(
                                ["--extractor", fallback]
                                if fallback in ("pymupdf4llm",
                                                "tesseract", "docling")
                                else []),
                            is_scanned=extractor_config.is_scanned,
                            avg_chars_per_page=(
                                extractor_config.avg_chars_per_page),
                            page_count=extractor_config.page_count,
                        )
                        cmd = _build_cmd_for_extractor(
                            fallback, args.input_file,
                            output_md, images_dir, short_name,
                            args.no_images,
                        )
                        # MAJOR-QC-2: Update checkpoint so crash
                        # recovery knows which extractor was used.
                        checkpoint["extractor"] = fallback
                        checkpoint["fallback_chain"] = checkpoint.get(
                            "fallback_chain", []) + [current]
                        _fb_timeout = (1500 if fallback == "marker"
                                       else None)
                        exit_code = run_command(
                            cmd,
                            f"Step 1 (quality-gate fallback): "
                            f"{fallback}",
                            timeout=_fb_timeout,
                        )
                        # MAJOR-QC-1: Re-check quality of fallback
                        # output. A fallback that exits 0 but produces
                        # empty output must not pass silently.
                        if exit_code == 0 and output_md.exists():
                            _fb_gate_name = fallback.capitalize()
                            _fb_gate_critical = _extractor_quality_gate(
                                _fb_gate_name, output_md, args.input_file)
                            if _fb_gate_critical:
                                print(f"  Quality gate CRITICAL for "
                                      f"{fallback}: forcing next fallback")
                                exit_code = 1  # Continue fallback loop
                        current = fallback
                    if exit_code != 0:
                        _fail("Quality gate fallback failed - "
                              "all extractors exhausted",
                              checkpoint, checkpoint_path)

        if _is_pymupdf_cv:
            print(f"\n{'─' * 40}")
            print("Step 1b: Cross-Validation (pdfplumber)")
            print('─' * 40)

            # NOTE (Issue 4): This calls pymupdf4llm a second time with
            # page_chunks=True. The first call (in convert-paper.py) does
            # not expose its chunks to us. For documents under ~50 pages
            # the overhead is negligible. For large HTA PDFs (200+ pages)
            # this doubles extraction time. A future optimization could
            # have convert-paper.py write chunks to a sidecar JSON file.
            md_chunks = _get_page_chunks(args.input_file)
            if md_chunks:
                cross_val_flags = cross_validate_extraction(
                    args.input_file, md_chunks
                )
                if cross_val_flags:
                    print(f"  Flagged {len(cross_val_flags)} page(s) "
                          "with >5% word mismatch:")
                    for flag_item in cross_val_flags:
                        print(f"    Page {flag_item['page']}: "
                              f"{flag_item['completeness']:.1%} complete, "
                              f"sample missing: "
                              f"{flag_item['missing_sample'][:5]}")
                    checkpoint["cross_validation_flags"] = cross_val_flags

                    # ── MinerU cross-validation fallback (Issue 2) ──
                    # If a large fraction of pages are flagged, pymupdf4llm
                    # is producing garbage.  Auto-switch to MinerU.
                    _total_pages = len(md_chunks)
                    _flagged_pages = len(cross_val_flags)
                    _flag_rate = (
                        _flagged_pages / _total_pages
                        if _total_pages > 0 else 0.0
                    )

                    # ── m4 fix: Slide-based PDF detection ──
                    # Presentation-style PDFs naturally have high flag
                    # rates because pymupdf4llm and pdfplumber handle
                    # slide layouts very differently.  Detect this case
                    # and emit a specific informational message instead
                    # of treating it as extraction failure.
                    _slide_based = False
                    _slide_based_but_garbled = False
                    if _flag_rate >= 0.90:
                        _slide_based = _is_slide_based_pdf(
                            args.input_file, md_chunks)
                        if _slide_based:
                            # ── RC3 fix: Check pymupdf4llm output
                            # quality before skipping MinerU ──
                            # Slide-based PDFs with animation layers
                            # (e.g. PowerPoint exported to PDF) can
                            # produce garbled text that genuinely
                            # needs MinerU fallback.  Count garbled
                            # markers in the extracted markdown.
                            _garbled_count = 0
                            _GARBLED_THRESHOLD = 5
                            if output_md.exists():
                                try:
                                    _slide_md = output_md.read_text(
                                        encoding="utf-8")
                                    _garbled_count = _slide_md.count(
                                        "[OCR garbled")
                                except Exception:
                                    pass
                            if _garbled_count > _GARBLED_THRESHOLD:
                                _slide_based_but_garbled = True
                                print(
                                    f"  Slide-based PDF detected but "
                                    f"output has {_garbled_count} "
                                    f"garbled markers (threshold: "
                                    f"{_GARBLED_THRESHOLD}) — forcing "
                                    f"MinerU fallback")
                                checkpoint[
                                    "is_slide_based_pdf"
                                ] = True
                                checkpoint[
                                    "slide_based_garbled_override"
                                ] = True
                                checkpoint[
                                    "cross_val_flag_rate_note"
                                ] = (
                                    "Slide-based PDF with garbled "
                                    "output — MinerU forced"
                                )
                            else:
                                print(
                                    f"  EXPECTED: Slide-based PDF "
                                    f"detected. High cross-validation "
                                    f"flag rate ({_flag_rate:.0%}) is "
                                    f"expected for presentation-style "
                                    f"documents. "
                                    f"No quality impact."
                                    + (f" (garbled markers: "
                                       f"{_garbled_count})"
                                       if _garbled_count > 0
                                       else ""))
                                checkpoint[
                                    "is_slide_based_pdf"
                                ] = True
                                checkpoint[
                                    "cross_val_flag_rate_note"
                                ] = (
                                    "High flag rate expected for "
                                    "slide-based PDF"
                                )

                    if (_flag_rate >= MINERU_FALLBACK_THRESHOLD
                            and _total_pages >= MINERU_FALLBACK_MIN_PAGES
                            and _mineru_available()
                            and (not _slide_based
                                 or _slide_based_but_garbled
                                 or _flag_rate >= 0.90)):
                        _mineru_ok = _trigger_mineru_fallback(
                            input_file=args.input_file,
                            output_md=output_md,
                            images_dir=images_dir,
                            short_name=short_name,
                            checkpoint=checkpoint,
                            checkpoint_path=checkpoint_path,
                            cross_val_flag_rate=_flag_rate,
                            page_count=_total_pages,
                        )
                        if _mineru_ok:
                            # Update extractor_config so downstream code
                            # knows MinerU was used (manifest generation,
                            # issue log, registry).
                            extractor_config = ExtractorConfig(
                                extractor="mineru",
                                script=str(
                                    SCRIPTS_DIR / "convert-mineru.py"),
                                extra_args=[],
                                is_scanned=False,
                                avg_chars_per_page=(
                                    extractor_config.avg_chars_per_page),
                                page_count=_total_pages,
                            )
                            # Sanity check on MinerU output
                            # (We do NOT re-run pymupdf4llm cross-
                            # validation here -- that is the test that
                            # already failed.  Instead, verify the
                            # MinerU-produced .md exists, has content,
                            # and images were copied.)
                            print(f"\n{'─' * 40}")
                            print("Step 1b-RECHECK: MinerU Output "
                                  "Sanity Check")
                            print('─' * 40)
                            _sanity_ok = True
                            _sanity_issues = []
                            if not output_md.exists():
                                _sanity_issues.append(
                                    "output .md does not exist")
                                _sanity_ok = False
                            else:
                                _md_lines = output_md.read_text(
                                    encoding="utf-8"
                                ).splitlines()
                                if len(_md_lines) < 100:
                                    _sanity_issues.append(
                                        f"output .md has only "
                                        f"{len(_md_lines)} lines "
                                        f"(expected >= 100)")
                                    _sanity_ok = False
                                else:
                                    print(f"  Markdown: "
                                          f"{len(_md_lines)} lines")
                            if images_dir.exists():
                                _img_files = list(images_dir.iterdir())
                                if not _img_files:
                                    _sanity_issues.append(
                                        "images directory is empty")
                                    _sanity_ok = False
                                else:
                                    print(f"  Images: "
                                          f"{len(_img_files)} files")
                            else:
                                _sanity_issues.append(
                                    "images directory does not exist")
                                # Not necessarily fatal -- PDF may
                                # have no images
                                print("  Images: directory not found "
                                      "(may be expected for text-only "
                                      "PDFs)")
                            if _sanity_ok:
                                print("  MinerU fallback: sanity check "
                                      "PASS")
                                checkpoint[
                                    "mineru_sanity_check"
                                ] = "PASS"
                            else:
                                print("  MinerU fallback: sanity check "
                                      "WARN")
                                for _si in _sanity_issues:
                                    print(f"    - {_si}")
                                checkpoint[
                                    "mineru_sanity_check"
                                ] = "WARN"
                                checkpoint[
                                    "mineru_sanity_issues"
                                ] = _sanity_issues
                        else:
                            print("  MinerU fallback failed. "
                                  "Continuing with pymupdf4llm output.")
                    elif (_flag_rate >= MINERU_FALLBACK_THRESHOLD
                          and _total_pages < MINERU_FALLBACK_MIN_PAGES
                          and not _slide_based):
                        print(f"  MinerU fallback: flag rate "
                              f"{_flag_rate:.1%} exceeds threshold but "
                              f"document has only {_total_pages} pages "
                              f"(min {MINERU_FALLBACK_MIN_PAGES}). "
                              "Skipping.")
                    elif (_flag_rate >= MINERU_FALLBACK_THRESHOLD
                          and not _mineru_available()
                          and not _slide_based):
                        print(f"  MinerU fallback: flag rate "
                              f"{_flag_rate:.1%} exceeds threshold but "
                              "MinerU is not available. Skipping.")

                else:
                    print("  Cross-validation: PASS (no significant "
                          "mismatches)")
            else:
                print("  Cross-validation: SKIPPED "
                      "(could not get page chunks)")

        # ── Post-processing: text cleanup passes (F8/F9/F12/F15/m11) ──
        # Applied after extraction (Step 1) but before QC (Step 2).
        # Fixes CJK artifacts, garbled OCR blocks, euro symbols,
        # table caption prefixes, and LaTeX accent notation.
        if output_md.exists():
            try:
                _raw_md = output_md.read_text(encoding="utf-8")
                _cleaned_md = _post_process_markdown(_raw_md)
                if _cleaned_md != _raw_md:
                    output_md.write_text(_cleaned_md, encoding="utf-8")
                    print("  Post-processing: text cleanup applied "
                          "(F8/F9/F12/F15/m11)")
                else:
                    print("  Post-processing: no changes needed")
            except Exception as _pp_err:
                _warn(f"Post-processing failed (non-fatal): {_pp_err}")

        # ── F14: Inject document_domain into YAML frontmatter ──
        # Runs after convert-paper.py (or convert-office.py) has generated
        # the output .md.  Scans the full markdown text for domain keywords
        # and inserts document_domain into the existing YAML frontmatter.
        if output_md.exists():
            try:
                _f14_md = output_md.read_text(encoding="utf-8")
                _f14_domain, _f14_count, _f14_kws = (
                    _detect_document_domain(_f14_md))
                print(f"  F14 domain detection: {_f14_domain} "
                      f"({_f14_count} keyword(s): "
                      f"{', '.join(_f14_kws[:5])})")
                checkpoint["document_domain"] = _f14_domain
                checkpoint["domain_keyword_count"] = _f14_count
                # Inject into existing YAML frontmatter if present
                if _f14_md.startswith("---\n"):
                    _yaml_end = _f14_md.find("\n---", 4)
                    if _yaml_end != -1:
                        _yaml_block = _f14_md[4:_yaml_end]
                        # Only inject if not already present
                        if "document_domain:" not in _yaml_block:
                            # Insert after document_type line, or
                            # at end of YAML block
                            _insert_after = "document_type:"
                            _dt_pos = _yaml_block.find(_insert_after)
                            if _dt_pos != -1:
                                _dt_end = _yaml_block.find(
                                    "\n", _dt_pos)
                                if _dt_end != -1:
                                    _yaml_block = (
                                        _yaml_block[:_dt_end]
                                        + f"\ndocument_domain: "
                                        f"{_f14_domain}"
                                        + _yaml_block[_dt_end:]
                                    )
                                else:
                                    _yaml_block += (
                                        f"\ndocument_domain: "
                                        f"{_f14_domain}")
                            else:
                                _yaml_block += (
                                    f"\ndocument_domain: "
                                    f"{_f14_domain}")
                            _f14_md = (
                                "---\n"
                                + _yaml_block
                                + _f14_md[_yaml_end:]
                            )
                            output_md.write_text(
                                _f14_md, encoding="utf-8")
                            print(f"  F14: document_domain "
                                  f"'{_f14_domain}' injected into "
                                  f"YAML frontmatter")
                        else:
                            print(f"  F14: document_domain already "
                                  f"present in YAML")
                    else:
                        print("  F14: YAML frontmatter malformed "
                              "(no closing ---)")
                else:
                    print("  F14: No YAML frontmatter found "
                          "(domain logged to checkpoint only)")
            except Exception as _f14_err:
                _warn(f"F14 domain detection failed "
                      f"(non-fatal): {_f14_err}")

        # ── Fix 3.12: Image-Markdown Sync ──
        # After image extraction (Step 1), the manifest may list images but
        # the MD may still contain the "No images extracted." placeholder.
        # Sync the MD with the manifest by replacing the placeholder with
        # an image index table.  Must run BEFORE Step 2 QC gate to avoid
        # manifest/index count mismatches.
        if output_md.exists():
            # Build manifest candidate list (same logic as Step 3 but
            # needed here before Step 2)
            _sync_manifest_candidates = [
                images_dir / "image-manifest.json",
            ]
            # Office formats: manifest at output_md.parent
            _sync_input_stem = args.input_file.stem
            _sync_short = short_name
            _sync_manifest_candidates.append(
                output_md.parent / f"{_sync_input_stem}_manifest.json")
            if _sync_short != _sync_input_stem:
                _sync_manifest_candidates.append(
                    output_md.parent / f"{_sync_short}_manifest.json")

            _sync_manifest = None
            for _cand in _sync_manifest_candidates:
                if _cand.exists():
                    _sync_manifest = _cand
                    break

            if _sync_manifest is not None:
                try:
                    _synced = sync_images_to_md(output_md, _sync_manifest)
                    if _synced:
                        print("  Fix 3.12: Synced MD with manifest "
                              "(replaced 'No images extracted' "
                              "placeholder with image index)")
                    else:
                        print("  Fix 3.12: No sync needed "
                              "(placeholder absent or no images)")
                except Exception as _sync_err:
                    _warn(f"Fix 3.12 image-markdown sync failed "
                          f"(non-fatal): {_sync_err}")

        # ── Write checkpoint before QC ──
        try:
            checkpoint_path.parent.mkdir(parents=True, exist_ok=True)
            checkpoint_path.write_text(json.dumps(checkpoint, indent=2))
        except Exception as e:
            _warn(f"Could not write checkpoint: {e}")

        # ── Pre-QC: Early image index for MinerU path ──
        # generate_image_index() normally runs at Step 6c (after QC).
        # When MinerU fallback is used the companion *-image-index.md does
        # not exist yet when Step 2 (QC gate) runs, causing a FAIL on the
        # manifest-consistency check.  Fix: generate it now so the companion
        # file is present before the QC gate.
        if (fmt == "pdf"
                and extractor_config is not None
                and extractor_config.extractor == "mineru"
                and not args.no_images):
            try:
                print(f"\n{'─' * 40}")
                print("Step 1c: Early Image Index (MinerU pre-QC)")
                print('─' * 40)
                _early_index_meta = generate_image_index(
                    source_path=args.input_file,
                    output_md=output_md,
                    fmt=fmt,
                    target_dir=(args.target_dir.resolve()
                                if args.target_dir else None),
                    images_dir=images_dir,
                    extractor="mineru",
                )
                if _early_index_meta:
                    print("  Early image index: PASS")
                else:
                    print("  Early image index: WARN (returned None)")
            except Exception as _e:
                print(f"  Early image index: WARN ({_e})")
                # Non-fatal — Step 2 may still pass if no manifest exists

        # ── Step 2: Structural QC (GATE) ──
        cmd = [sys.executable, str(SCRIPTS_DIR / "qc-structural.py"),
               str(output_md)]
        exit_code = run_command(cmd, "Step 2: qc-structural.py (GATE)")

        if exit_code == 1:
            print("\nPIPELINE STOPPED: Step 2 (structural QC) FAILED")
            print(f"Fix issues and re-run: python3 "
                  f"{SCRIPTS_DIR / 'qc-structural.py'} {output_md}")
            checkpoint["current_state"] = "qc_structural_failed"
            try:
                checkpoint_path.write_text(
                    json.dumps(checkpoint, indent=2))
            except Exception:
                pass
            # R6: Write issue log for QC failure when --target-dir is set
            if args.target_dir and not args.dry_run:
                _qc_target_dir = args.target_dir.resolve()
                if not args.dry_run:
                    _qc_target_dir.mkdir(parents=True, exist_ok=True)
                _qc_extractor = "unknown"
                if extractor_config:
                    _qc_extractor = extractor_config.extractor
                elif fmt in ("docx", "pptx", "txt"):
                    _qc_extractor = f"convert-office-{fmt}"
                append_issue_log(
                    target_dir=_qc_target_dir,
                    source_file=args.input_file,
                    output_md=output_md,
                    extractor=_qc_extractor,
                    issue_type="QC_FAIL",
                    severity="CRITICAL",
                    details="Step 2 (structural QC) FAILED — pipeline stopped.",
                    action_taken="Pipeline aborted. Fix issues and re-run.",
                )
            sys.exit(1)
        elif exit_code == 2:
            print("\nWARNING: Step 2 passed with warnings (continuing)")
        else:
            print("\nStep 2: PASS")

        checkpoint["current_state"] = "qc_structural_complete"

        # ── Step 3: Prepare Image Analysis (if images exist) ──
        #
        # Manifest path discovery (Phase 2 fix):
        #   PDF / MinerU: convert-paper.py writes images_dir/image-manifest.json
        #   Office (DOCX/PPTX): convert-office.py writes {basename}_manifest.json
        #     in output_md.parent (the output directory, NOT the images sub-dir).
        #
        # Strategy: build a prioritised candidate list and use the first one that
        # exists.  This lets the same code handle PDF, MinerU, and office formats
        # without format-specific branches at the if-manifest_path.exists() check.
        # convert-office.py uses input_path.stem as-is for basename, so the
        # manifest is {input_stem}_manifest.json.  short_name may differ
        # (lowercased, spaces-to-dashes) so we check both.
        input_stem = args.input_file.stem  # raw stem, e.g. "My Presentation"
        manifest_path_candidates = [
            # 1. PDF / convert-paper.py canonical location
            images_dir / "image-manifest.json",
            # 2. Office formats primary: convert-office.py uses raw input stem
            output_md.parent / f"{input_stem}_manifest.json",
            # 3. Office formats alternate: in case short_name was transformed
            output_md.parent / f"{short_name}_manifest.json",
        ]
        # Remove duplicates while preserving order (e.g. when input_stem == short_name)
        seen: set = set()
        manifest_path_candidates = [
            p for p in manifest_path_candidates
            if not (str(p) in seen or seen.add(str(p)))  # type: ignore[func-returns-value]
        ]

        manifest_path = None
        for candidate in manifest_path_candidates:
            if candidate.exists():
                manifest_path = candidate
                if candidate != manifest_path_candidates[0]:
                    print(f"\nStep 3: Found office manifest at {candidate}")
                break

        # Fallback: use the PDF/standard path as the target even if it doesn't
        # exist yet (MinerU generation below may create it).
        if manifest_path is None:
            manifest_path = manifest_path_candidates[0]

        # Issue 5: If MinerU was the extractor, it writes images to
        # images/mineru/ without creating an image-manifest.json.
        # Generate a basic manifest from the MinerU images directory
        # so that step 3 and step 4 (IMAGE NOTEs) are not skipped.
        if (extractor_config
                and extractor_config.extractor == "mineru"
                and not manifest_path.exists()):
            mineru_images_dir = output_md.parent / "images" / "mineru"
            if mineru_images_dir.exists():
                mineru_imgs = (
                    list(mineru_images_dir.glob("*.png"))
                    + list(mineru_images_dir.glob("*.jpg"))
                    + list(mineru_images_dir.glob("*.jpeg"))
                )
                if mineru_imgs:
                    manifest_data = {
                        "source": "mineru",
                        "md_file": str(output_md),
                        "images_dir": str(mineru_images_dir),
                        "image_count": len(mineru_imgs),
                        "generated": datetime.now().isoformat(),
                        "images": []
                    }
                    for idx, img_path in enumerate(sorted(mineru_imgs)):
                        manifest_data["images"].append({
                            "index": idx,
                            "figure_num": idx + 1,
                            "filename": img_path.name,
                            "path": str(img_path),
                            "file_path": str(img_path),
                            "page": None,
                            "width": 0,
                            "height": 0,
                            "type_guess": "unknown",
                            "section_context": "",
                            "detected_caption": None,
                        })
                    # Write manifest to the standard images_dir
                    images_dir.mkdir(parents=True, exist_ok=True)
                    manifest_path.write_text(
                        json.dumps(manifest_data, indent=2))
                    print(f"  Generated MinerU image manifest: "
                          f"{manifest_path} "
                          f"({len(mineru_imgs)} images)")

        if manifest_path.exists():
            cmd = [
                sys.executable,
                str(SCRIPTS_DIR / "prepare-image-analysis.py"),
                str(output_md),
                "--manifest", str(manifest_path)
            ]
            exit_code = run_command(
                cmd, "Step 3: prepare-image-analysis.py",
                allow_failure=True
            )
            if exit_code != 0:
                print("\nWARNING: Step 3 (prepare-analysis) failed. "
                      "IMAGE NOTE generation can proceed without it.")
        else:
            print("\nStep 3: SKIP (no images to analyze)")

        # ── Step 6a: Extract Numbers (PDF only, unless skipped) ──
        if fmt == "pdf" and not args.skip_numbers:
            cmd = [
                sys.executable,
                str(SCRIPTS_DIR / "extract-numbers.py"),
                str(args.input_file),
                str(output_md),
            ]
            exit_code = run_command(
                cmd, "Step 6a: extract-numbers.py (PDF only)",
                allow_failure=True
            )
            if exit_code != 0:
                print("\nWARNING: Number extraction failed. "
                      "QC content fidelity can proceed without it.")
        else:
            print(f"\nStep 6a: SKIP "
                  f"(format={fmt}, skip_numbers={args.skip_numbers})")

        # ── Step 6c: Image Index Generation (R19) ──
        # Runs AFTER all conversion/QC steps, BEFORE organization.
        # Generates per-file image manifest alongside the .md file.
        # Result is stored in _image_index_meta for R21 registry integration.
        _image_index_meta = None
        if args.no_images:
            print(f"\nStep 6c: SKIP (--no-images flag set)")
        else:
            try:
                _image_index_meta = generate_image_index(
                    source_path=args.input_file,
                    output_md=output_md,
                    fmt=fmt,
                    target_dir=(args.target_dir.resolve()
                                if args.target_dir else None),
                    images_dir=images_dir,  # BUG-1 fix: pass slug subdir
                    extractor=(extractor_config.extractor
                               if extractor_config else None),
                )
            except Exception as e:
                _warn(f"Image index generation failed: {e}")
                check_for_known_failures(str(e), context="Step 6c: Image Index Generation")
                # Non-fatal: conversion already succeeded

        # R19 FIX: For DOCX/PPTX, generate_image_index() returns None
        # because convert-office.py handles it.  Use the metadata we
        # discovered earlier (after the subprocess call) as fallback.
        if (_image_index_meta is None and fmt in ("docx", "pptx")
                and not args.no_images):
            _image_index_meta = _office_image_index_meta

        # ── m3: Apply overrides to PPTX/DOCX image index files ──
        # For office formats, convert-office.py already wrote the image
        # index file.  We apply overrides by rewriting the file on disk.
        if (fmt in ("docx", "pptx")
                and _image_index_meta is not None
                and _image_index_meta.get("image_index_path")):
            _office_idx_path = Path(
                _image_index_meta["image_index_path"])
            if _office_idx_path.exists():
                _m3_office_overrides = _load_image_index_overrides(
                    args.input_file,
                    target_dir=(args.target_dir.resolve()
                                if args.target_dir else None),
                )
                if _m3_office_overrides is not None:
                    _m3_office_count = _apply_overrides_to_image_index_file(
                        _office_idx_path,
                        _m3_office_overrides,
                        args.input_file.name,
                    )
                    if _m3_office_count > 0:
                        # Update metadata counts to reflect overrides
                        # Re-parse the rewritten file for accurate counts
                        try:
                            _reparse = _office_idx_path.read_text(
                                encoding="utf-8")
                            for _rline in _reparse.splitlines():
                                if _rline.startswith(
                                        "Estimated substantive"):
                                    _rval = _rline.split(":")[1].strip()
                                    _image_index_meta[
                                        "substantive_images"] = int(
                                        _rval.split()[0])
                                    _image_index_meta[
                                        "has_testable_images"] = (
                                        int(_rval.split()[0]) > 0)
                        except Exception:
                            pass  # Non-fatal

        # ── Step 3b: Re-run prepare-image-analysis if Step 3 was skipped ──
        # When extractor is marker (or any extractor that defers image
        # manifest creation to Step 6c), Step 3 runs before the manifest
        # exists and is skipped.  Step 6c then creates image-manifest.json
        # and renders vector pages.  Without this re-run, no
        # analysis-manifest.json is created and Step 4 (AI descriptions)
        # is blocked.
        #
        # Conditions (ALL must be true):
        #   1. image-manifest.json now exists (created/updated by Step 6c)
        #   2. analysis-manifest.json does NOT yet exist
        #   3. Images are not disabled (--no-images)
        if not args.no_images and manifest_path.exists():
            # Check all candidate locations for analysis-manifest.json
            _step3b_analysis_candidates = [
                images_dir / "analysis-manifest.json",
            ]
            if fmt != "pdf":
                _step3b_office_images = (
                    output_md.parent / f"{short_name}_images")
                _step3b_analysis_candidates.append(
                    _step3b_office_images / "analysis-manifest.json")
                _step3b_stem_images = (
                    output_md.parent / f"{args.input_file.stem}_images")
                if _step3b_stem_images != _step3b_office_images:
                    _step3b_analysis_candidates.append(
                        _step3b_stem_images / "analysis-manifest.json")
            _step3b_has_analysis = any(
                p.exists() for p in _step3b_analysis_candidates)

            if not _step3b_has_analysis:
                print("\nStep 3b: Re-running prepare-image-analysis "
                      "(manifest created by Step 6c)")
                cmd = [
                    sys.executable,
                    str(SCRIPTS_DIR / "prepare-image-analysis.py"),
                    str(output_md),
                    "--manifest", str(manifest_path)
                ]
                exit_code = run_command(
                    cmd, "Step 3b: prepare-image-analysis.py (post-6c)",
                    allow_failure=True
                )
                if exit_code != 0:
                    print("\nWARNING: Step 3b (prepare-analysis re-run) "
                          "failed. IMAGE NOTE generation can proceed "
                          "without it.")
                else:
                    print("Step 3b: DONE")
            else:
                print("\nStep 3b: SKIP "
                      "(analysis-manifest.json already exists)")
        elif not args.no_images:
            print("\nStep 3b: SKIP "
                  "(no image-manifest.json from Step 6c)")

        # ── Update Registry ──
        # PDF: always write here (run-pipeline.py is the sole orchestrator for PDF).
        # Office (DOCX/PPTX/TXT): convert-office.py writes its own registry entry
        # on success. We write a second entry here as a safety net, using the same
        # update_registry() function (which deduplicates by source_hash). Double-write
        # is harmless: dedup removes the older entry and keeps the newest.
        # XLSX is not routed here (early exit above); this branch never sees xlsx.
        if fmt == "pdf":
            # source_hash computed early (before Step 8 file move)
            extractor_used = (extractor_config.extractor
                              if extractor_config else "pymupdf4llm")
            try:
                update_registry(args.input_file, output_md,
                                source_hash, extractor_used,
                                image_index_meta=_image_index_meta)
            except Exception as e:
                _warn(f"Could not update registry: {e}")
                # Do not exit; conversion already succeeded
        elif fmt in ("docx", "pptx", "txt"):
            # Office formats: write registry entry so the hook can find the .md
            # by SHA-256 even when the .md is in a different directory than the
            # source file (the primary correctness gap that Phase 4 fixes).
            # source_hash computed early (before Step 8 file move)
            extractor_used = f"convert-office-{fmt}"
            try:
                update_registry(args.input_file, output_md,
                                source_hash, extractor_used,
                                image_index_meta=_image_index_meta)
            except Exception as e:
                _warn(f"Could not update registry for {fmt.upper()}: {e}")
                # Do not exit; conversion already succeeded

        # ── Final checkpoint ──
        checkpoint["current_state"] = "python_pipeline_complete"
        checkpoint["completed_at"] = datetime.now().isoformat()
        try:
            checkpoint_path.write_text(json.dumps(checkpoint, indent=2))
        except Exception:
            pass

        # ── Report Next Steps ──
        print("\n" + "=" * 60)
        print("PYTHON PIPELINE COMPLETE")
        print("=" * 60)
        print(f"MD file:    {output_md}")
        print(f"Images:     {images_dir}")
        if fmt == "pdf" and extractor_config:
            print(f"Extractor:  {extractor_config.extractor}")
            print(f"Scanned:    {extractor_config.is_scanned}")
        if cross_val_flags:
            print(f"Cross-val:  {len(cross_val_flags)} page(s) flagged")
        if manifest_path.exists():
            print(f"Manifest:   {manifest_path}")
            # Determine where prepare-image-analysis.py wrote its output.
            # It writes analysis-manifest.json into manifest["images_dir"].
            # For PDF this is images_dir; for office it is {basename}_images/.
            # Try both locations so the report is correct regardless of format.
            analysis_manifest_candidates = [
                images_dir / "analysis-manifest.json",
            ]
            # For office: images_dir in manifest = output_md.parent/{basename}_images
            if fmt != "pdf":
                office_images_dir = output_md.parent / f"{short_name}_images"
                analysis_manifest_candidates.append(
                    office_images_dir / "analysis-manifest.json"
                )
                # Also try with input stem in case short_name differs
                stem_images_dir = (
                    output_md.parent / f"{args.input_file.stem}_images"
                )
                if stem_images_dir != office_images_dir:
                    analysis_manifest_candidates.append(
                        stem_images_dir / "analysis-manifest.json"
                    )
            analysis_manifest = next(
                (p for p in analysis_manifest_candidates if p.exists()), None
            )
            if analysis_manifest:
                print(f"Analysis:   {analysis_manifest}")

        context_summary = output_md.parent / "context-summary.json"
        if context_summary.exists():
            print(f"Context:    {context_summary}")

        print(f"Checkpoint: {checkpoint_path}")

        print()
        print("NEXT: Run Claude subagent steps:")
        print(f"  4. generate-image-notes.md (Claude) - {output_md}")
        print(f"  5. validate-image-notes.py (Python) - {output_md}")
        if fmt == "pdf":
            number_diff = output_md.parent / f"{output_md.stem}-number-diff-report.json"
            if number_diff.exists():
                print(f"  6b. qc-content-fidelity.md (Claude) - "
                      f"{output_md}")
                print(f"     (uses {number_diff})")
            else:
                print(f"  6b. qc-content-fidelity.md (Claude) - "
                      f"{output_md} (fallback mode)")
            print(f"  7. qc-final-review.md (Claude) - {output_md}")
        else:
            print(f"  7. qc-final-review.md (Claude) - {output_md}")
            print(f"  (Content fidelity check skipped for "
                  f"{fmt.upper()})")
        print("=" * 60)

    # ── m2: Agent Descriptions File (when --agent-descriptions is set) ──
    # Generates a structured prompt file with everything needed to
    # describe substantive images using generate-image-notes.md.
    # Runs after conversion + image index are complete but before
    # organization.  Requires image index + analysis manifest to exist.
    if args.agent_descriptions:
        _m2_desc_path = None
        if (_image_index_meta is not None
                and _image_index_meta.get("image_index_path")):
            _m2_idx_path = Path(_image_index_meta["image_index_path"])
            if _m2_idx_path.exists():
                try:
                    _m2_desc_path = generate_agent_descriptions_file(
                        output_md=output_md,
                        image_index_path=_m2_idx_path,
                        images_dir=images_dir,
                        fmt=fmt,
                        short_name=short_name,
                        input_stem=args.input_file.stem,
                    )
                except Exception as e:
                    _warn(f"m2: Agent descriptions generation failed: {e}")
                    check_for_known_failures(str(e), context="m2: Agent Descriptions")
            else:
                _warn(f"m2: Image index not found at {_m2_idx_path}. "
                      "Skipping agent descriptions.")
        else:
            _warn("m2: No image index metadata available. "
                  "Run the pipeline first to generate an image index, "
                  "then retry with --agent-descriptions.")

        if _m2_desc_path:
            print(f"\nm2: Agent descriptions file: {_m2_desc_path}")

    # ── v3.1 Organization Phase ───────────────────────────────────────────
    # R16: ALL organization steps are gated behind args.target_dir.
    # When --target-dir is NOT provided, nothing below this comment runs.
    # This is the compatibility guard that ensures v3.0 behavior is preserved
    # for all callers that do not pass --target-dir.
    #
    # The following steps are implemented in later waves and will be filled
    # in here. The guard structure is established now so that all future
    # wave implementations slot in without restructuring main():
    #
    #   Step 7:  Pre-deletion verification (R5)        [Wave 2]
    #   Step 8:  Move source → _originals/ (R1)        [Wave 2]
    #   Step 9:  Place .md, images, manifest (R2, R3)  [Wave 2]
    #   Step 10: Cleanup intermediate files (R4)        [Wave 2]
    #   Step 11: Update registry organized paths (R10)  [Wave 3]
    #   Step 12: Write visual report to disk (R7)       [Wave 4]
    #   Step 13: Append to issue log (R6)               [Wave 4]
    #
    # BACKWARD COMPATIBILITY CONTRACT (R16):
    #   - No files are created or moved when args.target_dir is None
    #   - CONVERSION-ISSUES.md is NOT created without --target-dir
    #   - PIPELINE-REPORT-*.md is NOT written without --target-dir
    #   - _originals/ directory is NOT created without --target-dir
    #   - All v3.0 flags remain unchanged and fully functional
    if args.target_dir:
        # ── v3.1 Organization Phase (R15: execution order) ──────────────
        # R15 contract: Steps 7-13 run ONLY after all v3.0 steps above
        # have succeeded or passed with warnings. If we reach this point,
        # conversion + QC are complete.
        target_dir = args.target_dir.resolve()
        if not args.dry_run:
            target_dir.mkdir(parents=True, exist_ok=True)
        input_stem = args.input_file.stem

        # source_hash already computed early (before Step 8 file move)

        # ── Tracking lists for R7 visual report ────────────────────────
        _report_move_actions = []
        _report_place_actions = []
        _report_issues = []

        # ── R12: Registry-Aware Duplicate Handling ─────────────────────
        # Check registry BEFORE conversion/organization starts.
        # Same hash + same target-dir → skip, report ALREADY CONVERTED.
        # Same hash + different target-dir → proceed normally.
        # --force bypasses duplicate check entirely.
        if not args.force and not args.organize_only:
            dup_entry = check_registry_duplicate(source_hash, target_dir)
            if dup_entry is not None:
                print("\n" + "─" * 40)
                print("DUPLICATE DETECTED (R12)")
                print("─" * 40)
                print(f"  SHA-256: {source_hash.removeprefix('sha256:')}")
                print(f"  Previous conversion: "
                      f"{dup_entry.get('converted_at', 'unknown')}")
                print(f"  Previous output: "
                      f"{dup_entry.get('output_path', 'unknown')}")
                prev_target = dup_entry.get('target_dir', '')
                if prev_target:
                    print(f"  Previous target-dir: {prev_target}")
                print("  Status: ALREADY CONVERTED")
                print("  Use --force to re-convert.")
                print("─" * 40)
                # R12: Update organized_at timestamp even on duplicate skip
                if not args.dry_run:
                    try:
                        _dup_extractor = dup_entry.get(
                            "extractor_used", "unknown")
                        _dup_output = Path(
                            dup_entry.get("output_path", str(output_md)))
                        _dup_images = None
                        _dup_img_path = dup_entry.get(
                            "organized_images_path", "")
                        if _dup_img_path:
                            _dup_images = Path(_dup_img_path)
                        update_registry_organized(
                            source_hash=source_hash,
                            source_file=args.input_file,
                            output_md=_dup_output,
                            target_dir=target_dir,
                            extractor=_dup_extractor,
                            images_dir=_dup_images,
                        )
                        print("  organized_at timestamp updated.")
                    except Exception as e:
                        _warn(f"Could not update organized_at: {e}")
                sys.exit(0)

        # ── R14: Idempotent / Skip-if-Done Check ───────────────────────
        # Must run FIRST. If target dir already has a properly organized
        # copy with matching SHA-256, skip all organize steps.
        if not args.force and check_already_organized(args.input_file, target_dir):
            # Check if .md is also already in place
            target_md = target_dir / f"{input_stem}.md"
            if target_md.exists():
                print("\n" + "─" * 40)
                print("ORGANIZATION: ALREADY DONE (idempotent skip)")
                print("─" * 40)
                print(f"  Source already in: "
                      f"{target_dir / ORIGINALS_SUBDIR / args.input_file.name}")
                print(f"  Output already at: {target_md}")
                if args.dry_run:
                    print("  (DRY RUN) No changes needed.")
                print("=" * 60)
                sys.exit(0)

        # ── R5: Pre-Deletion Verification ──────────────────────────────
        # Verify converted output exists and is valid BEFORE any moves.
        # The .md may still be at its extraction location (not yet moved).
        print(f"\n{'─' * 40}")
        print("Step 7: Pre-Deletion Verification (R5)")
        print("─" * 40)

        verification_passed, v_warnings, v_errors, _has_frontmatter = \
            verify_conversion_output(output_md)

        if v_warnings:
            for w in v_warnings:
                print(f"  ⚠ {w}")
                _report_issues.append({
                    "severity": "WARNING",
                    "issue_type": "VERIFICATION_FAIL",
                    "details": w,
                })
        if v_errors:
            for e in v_errors:
                print(f"  ✗ {e}")
                _report_issues.append({
                    "severity": "CRITICAL",
                    "issue_type": "VERIFICATION_FAIL",
                    "details": e,
                })
            print("\nORGANIZATION ABORTED: Verification failed (R5).")
            print("  Conversion output is missing or empty.")
            print("  Source file remains at original location.")
            sys.exit(2)

        md_size = output_md.stat().st_size
        print(f"  ✓ Output .md exists and non-empty ({md_size:,} bytes)")
        if not v_warnings:
            print("  ✓ YAML frontmatter present")

        # ── R1: Auto-Organize Originals ────────────────────────────────
        # Move source file to [target-dir]/_originals/[filename]
        print(f"\n{'─' * 40}")
        print("Step 8: Move Source → _originals/ (R1)")
        print("─" * 40)

        originals_dir = target_dir / ORIGINALS_SUBDIR
        original_dest = originals_dir / args.input_file.name
        # Track the actual source location after R1 move for issue log / registry
        _source_for_logging = args.input_file

        if original_dest.exists():
            # R14: source already at destination
            print(f"  SKIP: Source already at {original_dest}")
            _source_for_logging = original_dest
            _report_move_actions.append(
                ("SKIP", f"{args.input_file.name} already in {ORIGINALS_SUBDIR}/"))
        elif args.dry_run:
            print(f"  (DRY RUN) Would move: {args.input_file}")
            print(f"         → {original_dest}")
            _report_move_actions.append(
                ("DRY RUN", f"Would move {args.input_file.name} → "
                 f"{ORIGINALS_SUBDIR}/"))
        else:
            try:
                atomic_move(args.input_file, original_dest)
                _source_for_logging = original_dest
                print(f"  MOVED: {args.input_file.name}")
                print(f"    FROM: {args.input_file.parent}")
                print(f"    TO:   {originals_dir}")
                _report_move_actions.append(
                    ("MOVED", f"{args.input_file.name}\n"
                     f"         FROM: {_truncate_path(str(args.input_file.parent))}\n"
                     f"         TO:   {ORIGINALS_SUBDIR}/"))
            except Exception as e:
                print(f"  ✗ MOVE FAILED: {e}")
                print("  Source file remains at original location.")
                _report_issues.append({
                    "severity": "CRITICAL",
                    "issue_type": "MOVE_FAIL",
                    "details": f"Move source failed: {e}",
                })
                # R6: Write issue log before exit so CRITICAL entry reaches disk
                _ext_for_log = "unknown"
                if extractor_config:
                    _ext_for_log = extractor_config.extractor
                elif fmt in ("docx", "pptx", "txt"):
                    _ext_for_log = f"convert-office-{fmt}"
                for issue in _report_issues:
                    append_issue_log(
                        target_dir=target_dir,
                        source_file=args.input_file,
                        output_md=output_md,
                        extractor=_ext_for_log,
                        issue_type=issue["issue_type"],
                        severity=issue["severity"],
                        details=issue["details"],
                        action_taken="Pipeline aborted (move failure).",
                    )
                sys.exit(3)

        # ── R2: Auto-Place Conversions ─────────────────────────────────
        # Move .md to [target-dir]/[input_stem].md
        print(f"\n{'─' * 40}")
        print("Step 9a: Place .md in Target Dir (R2)")
        print("─" * 40)

        target_md = target_dir / f"{input_stem}.md"

        if output_md.resolve() == target_md.resolve():
            print(f"  SKIP: .md already at {target_md}")
            _report_place_actions.append(
                ("SKIP", f"{target_md.name} already in place"))
        elif target_md.exists():
            # R14: .md already exists at target (idempotent)
            print(f"  SKIP: .md already exists at {target_md}")
            _report_place_actions.append(
                ("SKIP", f"{target_md.name} already exists"))
        elif args.dry_run:
            print(f"  (DRY RUN) Would place: {output_md.name}")
            print(f"         → {target_md}")
            _report_place_actions.append(
                ("DRY RUN", f"Would place {output_md.name} → {target_dir}"))
        else:
            try:
                atomic_move(output_md, target_md)
                output_md = target_md  # update reference for later steps
                print(f"  PLACED: {target_md.name}")
                print(f"    AT: {target_dir}")
                _report_place_actions.append(
                    ("PLACED", f"{target_md.name} AT: "
                     f"{_truncate_path(str(target_dir))}"))
            except Exception as e:
                print(f"  ✗ PLACEMENT FAILED: {e}")
                _report_issues.append({
                    "severity": "CRITICAL",
                    "issue_type": "MOVE_FAIL",
                    "details": f"Place .md failed: {e}",
                })
                # R6: Write issue log before exit so CRITICAL entry reaches disk
                _ext_for_log = "unknown"
                if extractor_config:
                    _ext_for_log = extractor_config.extractor
                elif fmt in ("docx", "pptx", "txt"):
                    _ext_for_log = f"convert-office-{fmt}"
                for issue in _report_issues:
                    append_issue_log(
                        target_dir=target_dir,
                        source_file=_source_for_logging,
                        output_md=output_md,
                        extractor=_ext_for_log,
                        issue_type=issue["issue_type"],
                        severity=issue["severity"],
                        details=issue["details"],
                        action_taken="Pipeline aborted (placement failure).",
                    )
                sys.exit(3)

        # ── R3: Auto-Place Image Directories ───────────────────────────
        # Move [input_stem]_images/ and [input_stem]_manifest.json
        print(f"\n{'─' * 40}")
        print("Step 9b: Place Images + Manifest (R3)")
        print("─" * 40)

        # Locate the source images directory (may be in various locations)
        # convert-office.py writes: output_dir/{input_stem}_images/
        # convert-paper.py writes: images_dir (from -i flag or default)
        source_images_candidates = [
            output_md.parent / f"{input_stem}_images",
            images_dir,
            # Original output dir (before .md was moved)
            args.input_file.parent / f"{input_stem}_images",
        ]
        # Deduplicate while preserving order
        seen_img: set = set()
        source_images_candidates = [
            p for p in source_images_candidates
            if not (str(p.resolve()) in seen_img
                    or seen_img.add(str(p.resolve())))  # type: ignore[func-returns-value]
        ]

        source_images_dir = None
        for candidate in source_images_candidates:
            if candidate.exists() and candidate.is_dir():
                source_images_dir = candidate
                break

        target_images_dir = target_dir / f"{input_stem}_images"

        if source_images_dir is not None:
            img_actions = move_images_dir(
                source_images_dir, target_images_dir,
                target_md if target_md.exists() else output_md,
                dry_run=args.dry_run,
            )
            for action_type, desc in img_actions:
                print(f"  {action_type}: {desc}")
                _report_place_actions.append((action_type, desc))
            # Fix 3.9/M8: Rewrite image paths in MD file after move
            if not args.dry_run and target_md.exists():
                _md_content_39 = target_md.read_text(encoding='utf-8')
                _old_img_dir = source_images_dir.name
                _new_img_dir = target_images_dir.name
                if _old_img_dir != _new_img_dir:
                    _md_changed_39 = False
                    if f']({_old_img_dir}/' in _md_content_39:
                        _md_content_39 = _md_content_39.replace(
                            f']({_old_img_dir}/',
                            f']({_new_img_dir}/')
                        _md_changed_39 = True
                    if f'src="{_old_img_dir}/' in _md_content_39:
                        _md_content_39 = _md_content_39.replace(
                            f'src="{_old_img_dir}/',
                            f'src="{_new_img_dir}/')
                        _md_changed_39 = True
                    # MAJOR-3: Also rewrite paths inside IMAGE HTML comments
                    if f'IMAGE: {_old_img_dir}/' in _md_content_39:
                        _md_content_39 = _md_content_39.replace(
                            f'IMAGE: {_old_img_dir}/',
                            f'IMAGE: {_new_img_dir}/')
                        _md_changed_39 = True
                    if _md_changed_39:
                        target_md.write_text(
                            _md_content_39, encoding='utf-8')
                        print(f"  REWRITTEN: MD image paths updated "
                              f"({_old_img_dir} -> {_new_img_dir})")
        elif args.no_images:
            print("  SKIP: --no-images flag set, no images to place")
        else:
            print("  SKIP: No images directory found")

        # ── Helper: rewrite manifest JSON paths ──────────────────────
        # Defined here so it is available to both the analysis-manifest
        # block below (line ~4658) and the source-manifest block further
        # down (lines ~4851-4882).  Must come before its first call.
        def _rewrite_manifest_paths(mf_path: Path, old_base: str,
                                    new_base: str) -> bool:
            """Update images_dir / md_file / per-image file_path inside a manifest JSON.

            Rewrites:
              - top-level images_dir
              - top-level md_file
              - every images[N].file_path entry

            Returns True if the file was changed, False otherwise.
            old_base and new_base are resolved directory path strings.
            """
            if old_base == new_base:
                return False
            try:
                import json as _json2
                _mf_data = _json2.loads(
                    mf_path.read_text(encoding="utf-8"))
                _mf_changed = False
                for _key in ("images_dir", "md_file"):
                    _val = _mf_data.get(_key, "")
                    if _val and old_base in _val:
                        _mf_data[_key] = _val.replace(old_base, new_base)
                        _mf_changed = True
                # Rewrite per-image file_path entries
                for _img in _mf_data.get("images", []):
                    _fp = _img.get("file_path", "")
                    if _fp and old_base in _fp:
                        _img["file_path"] = _fp.replace(old_base, new_base)
                        _mf_changed = True
                    # Rewrite nested template_skeleton.file if present
                    _ts = _img.get("template_skeleton", {})
                    _ts_file = _ts.get("file", "")
                    if _ts_file and old_base in _ts_file:
                        _ts["file"] = _ts_file.replace(old_base, new_base)
                        _mf_changed = True
                if _mf_changed:
                    mf_path.write_text(
                        _json2.dumps(_mf_data, indent=2,
                                     ensure_ascii=False),
                        encoding="utf-8")
                return _mf_changed
            except Exception as _rw_e:
                _warn(f"Could not rewrite paths in manifest: {_rw_e}")
                return False

        # ── Rewrite analysis-manifest.json paths if present ──────────
        # prepare-image-analysis.py writes analysis-manifest.json inside
        # the _images/ directory.  It records absolute paths baked in at
        # write time (source dir).  After move_images_dir() relocates the
        # images directory to target_dir, rewrite any stale paths.
        if not args.dry_run and target_images_dir.exists():
            _analysis_manifest = target_images_dir / "analysis-manifest.json"
            if _analysis_manifest.exists():
                _am_old = str(args.input_file.parent.resolve())
                _am_new = str(target_dir.resolve())
                if _rewrite_manifest_paths(_analysis_manifest,
                                           _am_old, _am_new):
                    print(f"  REWRITTEN: analysis-manifest.json paths updated")

        # ── C1/C2 FIX: Copy vector renders to output images dir ──────
        # Vector renders are saved to images_dir (images/<slug>/) during
        # Step 6c, but Step 9b may not move them to target_images_dir
        # because: (a) on re-runs, target_images_dir already exists from
        # a previous run so move_images_dir SKIPs, (b) source_images_dir
        # may resolve to a different candidate than images_dir.
        # Fix: explicitly copy any vector renders from images_dir to
        # target_images_dir after Step 9b completes.
        if (fmt == "pdf" and not args.no_images
                and images_dir.exists()):
            _vr_files = sorted(images_dir.glob("*-vector-render.png"))
            if _vr_files:
                if not args.dry_run:
                    target_images_dir.mkdir(parents=True, exist_ok=True)
                _vr_copied = 0
                for _vr_file in _vr_files:
                    _vr_dest = target_images_dir / _vr_file.name
                    if _vr_dest.exists():
                        continue  # Already there (first-run move)
                    if args.dry_run:
                        print(f"  (DRY RUN) Would copy vector render: "
                              f"{_vr_file.name}")
                        _vr_copied += 1
                    else:
                        try:
                            shutil.copy2(str(_vr_file), str(_vr_dest))
                            _ensure_max_dimension(_vr_dest)
                            _vr_copied += 1
                        except Exception as _vr_e:
                            _warn(f"Could not copy vector render "
                                  f"{_vr_file.name}: {_vr_e}")
                if _vr_copied > 0:
                    print(f"  C1/C2 FIX: Copied {_vr_copied} vector "
                          f"render(s) to {target_images_dir.name}")
                    _report_place_actions.append(
                        ("PLACED", f"{_vr_copied} vector render(s) → "
                         f"{target_images_dir.name}"))

                    # Update manifest in target_images_dir to include
                    # vector renders (merge from images_dir manifest).
                    _src_manifest = images_dir / "image-manifest.json"
                    _tgt_manifest = target_images_dir / "image-manifest.json"
                    if _src_manifest.exists() and not args.dry_run:
                        try:
                            _src_data = json.loads(
                                _src_manifest.read_text(encoding="utf-8"))
                            _src_vr_entries = [
                                img for img in _src_data.get("images", [])
                                if img.get("source") == "vector-render"
                            ]
                            if _src_vr_entries and _tgt_manifest.exists():
                                _tgt_data = json.loads(
                                    _tgt_manifest.read_text(
                                        encoding="utf-8"))
                                _existing = {
                                    img.get("filename")
                                    for img in _tgt_data.get("images", [])
                                }
                                _added_vr = 0
                                for _vr_entry in _src_vr_entries:
                                    if (_vr_entry.get("filename")
                                            not in _existing):
                                        # Update file_path to target dir
                                        _vr_entry["file_path"] = str(
                                            target_images_dir
                                            / _vr_entry["filename"])
                                        _tgt_data.setdefault(
                                            "images", []).append(_vr_entry)
                                        _added_vr += 1
                                if _added_vr > 0:
                                    _tgt_data["image_count"] = len(
                                        _tgt_data.get("images", []))
                                    _tgt_data["total_images"] = (
                                        _tgt_data["image_count"])
                                    _tgt_manifest.write_text(
                                        json.dumps(
                                            _tgt_data, indent=2,
                                            ensure_ascii=False) + "\n",
                                        encoding="utf-8")
                                    print(f"  C1/C2 FIX: Added "
                                          f"{_added_vr} vector render "
                                          f"entries to manifest")
                            elif _src_vr_entries and not _tgt_manifest.exists():
                                # No target manifest yet — copy source as-is
                                # but update file_path for ALL entries.
                                # M-2 FIX: The old code only updated
                                # file_path for vector render entries.
                                # Raster image entries retained stale paths
                                # pointing to _originals/images/<slug>/.
                                # Now we update every entry's file_path AND
                                # filter out entries whose files don't exist
                                # in the target directory.
                                _updated_images = []
                                for _img_entry in _src_data.get("images", []):
                                    _fname = _img_entry.get("filename", "")
                                    _tgt_file = target_images_dir / _fname
                                    if _tgt_file.exists() or args.dry_run:
                                        _img_entry["file_path"] = str(
                                            target_images_dir / _fname)
                                        _updated_images.append(_img_entry)
                                _src_data["images"] = _updated_images
                                _src_data["image_count"] = len(
                                    _updated_images)
                                _src_data["total_images"] = len(
                                    _updated_images)
                                _src_data["images_dir"] = str(
                                    target_images_dir)
                                _tgt_manifest.write_text(
                                    json.dumps(
                                        _src_data, indent=2,
                                        ensure_ascii=False) + "\n",
                                    encoding="utf-8")
                                print(f"  C1/C2 FIX: Created manifest "
                                      f"with {len(_updated_images)} "
                                      f"entries ({len(_src_vr_entries)} "
                                      f"vector renders)")
                        except Exception as _mf_e:
                            _warn(f"C1/C2: Could not update manifest "
                                  f"with vector renders: {_mf_e}")

        # Move manifest JSON
        source_manifest_candidates = [
            manifest_path,  # discovered earlier in Step 3
            output_md.parent / f"{input_stem}_manifest.json",
            args.input_file.parent / f"{input_stem}_manifest.json",
        ]
        seen_mf: set = set()
        source_manifest_candidates = [
            p for p in source_manifest_candidates
            if p is not None
            and not (str(p.resolve()) in seen_mf
                     or seen_mf.add(str(p.resolve())))  # type: ignore[func-returns-value]
        ]

        target_manifest = target_dir / f"{input_stem}_manifest.json"
        source_manifest = None
        for candidate in source_manifest_candidates:
            if candidate.exists() and candidate.is_file():
                source_manifest = candidate
                break

        if source_manifest is not None:
            if source_manifest.resolve() == target_manifest.resolve():
                print(f"  SKIP: Manifest already at {target_manifest}")
                _report_place_actions.append(
                    ("SKIP", f"{target_manifest.name} already in place"))
                # Still rewrite stale internal paths if needed
                if not args.dry_run:
                    _mf_old = str(source_manifest.parent.resolve())
                    _mf_new = str(target_dir.resolve())
                    if _rewrite_manifest_paths(target_manifest,
                                               _mf_old, _mf_new):
                        print(f"  REWRITTEN: Manifest "
                              f"images_dir/md_file paths updated")
            elif target_manifest.exists():
                print(f"  SKIP: Manifest already exists at {target_manifest}")
                _report_place_actions.append(
                    ("SKIP", f"{target_manifest.name} already exists"))
                # Still rewrite stale internal paths if needed
                if not args.dry_run:
                    _mf_old = str(source_manifest.parent.resolve())
                    _mf_new = str(target_dir.resolve())
                    if _rewrite_manifest_paths(target_manifest,
                                               _mf_old, _mf_new):
                        print(f"  REWRITTEN: Manifest "
                              f"images_dir/md_file paths updated")
            elif args.dry_run:
                print(f"  (DRY RUN) Would place manifest → {target_manifest}")
                _report_place_actions.append(
                    ("DRY RUN", f"Would place {target_manifest.name}"))
            else:
                try:
                    _old_manifest_base = str(source_manifest.parent.resolve())
                    atomic_move(source_manifest, target_manifest)
                    print(f"  PLACED: {target_manifest.name} → {target_dir}")
                    _report_place_actions.append(
                        ("PLACED", f"{target_manifest.name} → "
                         f"{_truncate_path(str(target_dir))}"))
                    # ── Rewrite stale images_dir / md_file paths ──
                    # The manifest records absolute paths that point to the
                    # pre-move directory.  Update them to the target_dir.
                    if _rewrite_manifest_paths(
                            target_manifest,
                            _old_manifest_base,
                            str(target_dir.resolve())):
                        print(f"  REWRITTEN: Manifest "
                              f"images_dir/md_file paths updated")
                except Exception as e:
                    print(f"  ⚠ Could not move manifest: {e}")
                    _report_place_actions.append(
                        ("WARNING", f"Could not move manifest: {e}"))
        else:
            print("  SKIP: No manifest file found")

        # ── R19: Move Image Index to Target Dir ──────────────────────
        # The image index file was generated in Step 6c alongside the .md.
        # If it exists and needs moving, move it now.
        try:
            _image_index_meta  # noqa: F841
        except NameError:
            _image_index_meta = None

        if _image_index_meta and _image_index_meta.get("image_index_path"):
            source_index = Path(_image_index_meta["image_index_path"])
            target_index = target_dir / source_index.name

            if source_index.exists() and source_index.resolve() != target_index.resolve():
                if args.dry_run:
                    print(f"  (DRY RUN) Would place image index → {target_index}")
                    _report_place_actions.append(
                        ("DRY RUN", f"Would place {source_index.name}"))
                else:
                    try:
                        atomic_move(source_index, target_index)
                        _image_index_meta["image_index_path"] = str(target_index)
                        print(f"  PLACED: {target_index.name} → {target_dir}")
                        _report_place_actions.append(
                            ("PLACED", f"{target_index.name} → "
                             f"{_truncate_path(str(target_dir))}"))
                        # ── Rewrite stale Source:/Converted: paths ──
                        # The image index records Source: and Converted:
                        # with the pre-move (_originals/) paths.  Replace
                        # the old directory prefix with target_dir.
                        try:
                            _idx_old_base = str(
                                source_index.parent.resolve())
                            _idx_new_base = str(target_dir.resolve())
                            if _idx_old_base != _idx_new_base:
                                _idx_content = target_index.read_text(
                                    encoding="utf-8")
                                _idx_content = _idx_content.replace(
                                    _idx_old_base, _idx_new_base)
                                target_index.write_text(
                                    _idx_content, encoding="utf-8")
                                print(f"  REWRITTEN: Image index "
                                      f"Source/Converted paths updated")
                        except Exception as _idx_rw_e:
                            _warn(f"Could not rewrite paths in "
                                  f"image index: {_idx_rw_e}")
                    except Exception as e:
                        print(f"  ⚠ Could not move image index: {e}")
                        _report_place_actions.append(
                            ("WARNING", f"Could not move image index: {e}"))
            elif source_index.exists():
                print(f"  SKIP: Image index already at {target_index}")
                _report_place_actions.append(
                    ("SKIP", f"{target_index.name} already in place"))

        # ── m2: Move agent descriptions file to target dir ──
        # If --agent-descriptions generated a file, move it alongside
        # the image index in the target directory.
        # _m2_desc_path is set in the agent_descriptions block above;
        # ensure it's defined when that block was not entered.
        try:
            _m2_desc_path
        except NameError:
            _m2_desc_path = None

        if (_m2_desc_path is not None
                and _m2_desc_path.exists()):
            _m2_source = _m2_desc_path
            _m2_target = target_dir / _m2_source.name
            if _m2_source.resolve() != _m2_target.resolve():
                if args.dry_run:
                    print(f"  (DRY RUN) Would place agent descriptions "
                          f"→ {_m2_target}")
                    _report_place_actions.append(
                        ("DRY RUN",
                         f"Would place {_m2_source.name}"))
                else:
                    try:
                        atomic_move(_m2_source, _m2_target)
                        _m2_desc_path = _m2_target
                        print(f"  PLACED: {_m2_target.name} "
                              f"→ {target_dir}")
                        _report_place_actions.append(
                            ("PLACED",
                             f"{_m2_target.name} → "
                             f"{_truncate_path(str(target_dir))}"))
                        # ── Rewrite stale paths inside the moved file ──
                        # The file was generated before the organization
                        # phase, so all absolute paths point to the
                        # pre-move source directory.  Replace them with
                        # the final target_dir paths.
                        try:
                            _old_base = str(
                                _m2_source.parent.resolve())
                            _new_base = str(target_dir.resolve())
                            if _old_base != _new_base:
                                _desc_content = _m2_target.read_text(
                                    encoding="utf-8")
                                _desc_content = _desc_content.replace(
                                    _old_base, _new_base)
                                _m2_target.write_text(
                                    _desc_content, encoding="utf-8")
                                print(f"  REWRITTEN: Internal paths "
                                      f"updated to {target_dir}")
                        except Exception as _rw_e:
                            _warn(f"m2: Could not rewrite paths "
                                  f"in agent descriptions: "
                                  f"{_rw_e}")
                    except Exception as e:
                        print(f"  ⚠ Could not move agent "
                              f"descriptions: {e}")
                        _report_place_actions.append(
                            ("WARNING",
                             f"Could not move agent descriptions: "
                             f"{e}"))

        # ── Issue-A: Move number-diff report and context-summary.json ──
        # extract-numbers.py writes {stem}-number-diff-report.json and
        # convert-paper.py writes context-summary.json to the source
        # directory (alongside the .md before it was moved).  Move them
        # to target_dir so all outputs are co-located.
        _source_dir_for_artifacts = args.input_file.parent
        _artifact_files = [
            (
                _source_dir_for_artifacts
                / f"{input_stem}-number-diff-report.json"
            ),
            _source_dir_for_artifacts / "context-summary.json",
        ]
        for _art_src in _artifact_files:
            if not _art_src.exists():
                continue
            _art_dst = target_dir / _art_src.name
            if _art_src.resolve() == _art_dst.resolve():
                continue
            if args.dry_run:
                print(f"  (DRY RUN) Would move {_art_src.name} "
                      f"→ {target_dir}")
                _report_place_actions.append(
                    ("DRY RUN",
                     f"Would place {_art_src.name} → "
                     f"{_truncate_path(str(target_dir))}"))
            else:
                try:
                    atomic_move(_art_src, _art_dst)
                    print(f"  PLACED: {_art_src.name} → {target_dir}")
                    _report_place_actions.append(
                        ("PLACED",
                         f"{_art_src.name} → "
                         f"{_truncate_path(str(target_dir))}"))
                except Exception as _art_e:
                    _warn(f"Could not move {_art_src.name}: {_art_e}")
                    _report_place_actions.append(
                        ("WARNING",
                         f"Could not move {_art_src.name}: {_art_e}"))

        # ── Fix 3.9/m20: Update YAML source_path in MD file ────────────
        # After all moves are complete, the source_path in the YAML
        # frontmatter still points to the original location.  Update it
        # to point to the new location in _originals/.
        if not args.dry_run and target_md.exists():
            try:
                _md_c_39 = target_md.read_text(encoding='utf-8')
                _yaml_m_39 = re.match(
                    r'(---\n)(.*?)(---\n)', _md_c_39, re.DOTALL)
                if _yaml_m_39:
                    _yb_39 = _yaml_m_39.group(2)
                    _new_src_39 = str(
                        target_dir / ORIGINALS_SUBDIR
                        / args.input_file.name)
                    _old_sp_match = re.search(
                        r'^source_path:\s*(.*)$',
                        _yb_39, re.MULTILINE)
                    if _old_sp_match:
                        _old_sp_val = _old_sp_match.group(1).strip().strip('"')
                        if _old_sp_val != _new_src_39:
                            # MINOR-5: Use re.escape to avoid backslash
                            # interpretation in replacement string
                            _escaped_src_39 = _new_src_39.replace(
                                '\\', '\\\\')
                            _yb_39 = re.sub(
                                r'(source_path:\s*).*',
                                rf'\1"{_escaped_src_39}"',
                                _yb_39
                            )
                            _md_c_39 = (
                                _yaml_m_39.group(1) + _yb_39
                                + _yaml_m_39.group(3)
                                + _md_c_39[_yaml_m_39.end():])
                            target_md.write_text(
                                _md_c_39, encoding='utf-8')
                            print(f"  REWRITTEN: YAML source_path updated "
                                  f"to {ORIGINALS_SUBDIR}/")
            except Exception as _sp_e:
                _warn(f"Fix 3.9/m20: Could not update YAML "
                      f"source_path: {_sp_e}")

        # ── R4: Cleanup Intermediate Files ─────────────────────────────
        print(f"\n{'─' * 40}")
        print("Step 10: Cleanup Intermediate Files (R4)")
        print("─" * 40)

        cleanup_actions = cleanup_intermediate_files(
            output_md=target_md if target_md.exists() else output_md,
            source_file=args.input_file,
            checkpoint_path=checkpoint_path,
            dry_run=args.dry_run,
        )

        # Also clean up source-dir artifacts left behind (R18 safety net):
        # Only clean artifacts that match THIS conversion's naming pattern.
        # CRITICAL: Do NOT glob /tmp/soffice-* (chart rendering cleans its own).
        source_dir = args.input_file.parent
        if source_dir.exists() and source_dir.resolve() != target_dir.resolve():
            # Images dir left in source dir (if R3 moved it out)
            leftover_images = source_dir / f"{input_stem}_images"
            if leftover_images.exists() and not any(leftover_images.iterdir()):
                if args.dry_run:
                    cleanup_actions.append(
                        ("DRY RUN", f"Would delete empty {leftover_images}"))
                else:
                    try:
                        shutil.rmtree(leftover_images)
                        cleanup_actions.append(
                            ("DELETED", str(leftover_images)))
                    except Exception as e:
                        cleanup_actions.append(
                            ("WARNING", f"Could not delete {leftover_images}: {e}"))

            # Manifest left in source dir
            leftover_manifest = source_dir / f"{input_stem}_manifest.json"
            if leftover_manifest.exists():
                if args.dry_run:
                    cleanup_actions.append(
                        ("DRY RUN", f"Would delete {leftover_manifest}"))
                else:
                    try:
                        leftover_manifest.unlink()
                        cleanup_actions.append(
                            ("DELETED", str(leftover_manifest)))
                    except Exception as e:
                        cleanup_actions.append(
                            ("WARNING",
                             f"Could not delete {leftover_manifest}: {e}"))

        if cleanup_actions:
            for action_type, desc in cleanup_actions:
                print(f"  {action_type}: {desc}")
        else:
            print("  No intermediate files to clean up.")

        # Determine the extractor name for registry and reporting
        _extractor_name = "unknown"
        if extractor_config:
            _extractor_name = extractor_config.extractor
        elif fmt in ("docx", "pptx", "txt"):
            _extractor_name = f"convert-office-{fmt}"

        # ── R11: Capture fallback extractor usage as WARNING ─────────
        # R6 spec: fallback extractor usage produces a WARNING entry.
        if (not args.organize_only
                and checkpoint.get("fallback_chain")):
            fallback_chain = checkpoint["fallback_chain"]
            fallback_detail = (
                f"Fallback extractor used. Chain: "
                f"{' → '.join(fallback_chain)} → {_extractor_name}"
            )
            _report_issues.append({
                "severity": "WARNING",
                "issue_type": "EXTRACTION_WARNING",
                "details": fallback_detail,
            })

        # ── R10 + R21: Update Registry with Organized Paths + Image Index ──
        print(f"\n{'─' * 40}")
        print("Step 11: Update Registry (R10 + R21)")
        print("─" * 40)

        # R21: Resolve image_index_meta for registry.
        # _image_index_meta is set during Step 6c (conversion phase).
        # If organize-only, try to discover existing image index.
        try:
            _image_index_meta  # noqa: F841 — test if bound
        except NameError:
            _image_index_meta = None

        # R21: If image index was generated but file was moved to target_dir,
        # update the path in the metadata to point to the new location.
        if _image_index_meta and _image_index_meta.get("image_index_path"):
            _idx_path = Path(_image_index_meta["image_index_path"])
            _target_idx = target_dir / _idx_path.name
            if _target_idx.exists():
                _image_index_meta["image_index_path"] = str(_target_idx)
            elif _idx_path.exists():
                _image_index_meta["image_index_path"] = str(_idx_path)

        if args.dry_run:
            print("  (DRY RUN) Would update registry with organized paths.")
            if _image_index_meta:
                print("  (DRY RUN) Would include R21 image index fields.")
        else:
            try:
                _org_images = (target_images_dir
                               if target_images_dir.exists() else None)
                update_registry_organized(
                    source_hash=source_hash,
                    source_file=_source_for_logging,
                    output_md=target_md if target_md.exists() else output_md,
                    target_dir=target_dir,
                    extractor=_extractor_name,
                    images_dir=_org_images,
                    image_index_meta=_image_index_meta,
                )
            except Exception as e:
                _warn(f"Could not update registry (R10/R21): {e}")
                _report_issues.append({
                    "severity": "WARNING",
                    "issue_type": "EXTRACTION_WARNING",
                    "details": f"Registry update failed: {e}",
                })

        # ── R6: Write issues to CONVERSION-ISSUES.md ────────────────
        # Only write if there are actual issues to log (R6 spec:
        # do not write "Conversion OK" entries).
        if _report_issues and not args.dry_run:
            print(f"\n{'─' * 40}")
            print("Step 13: Write Issue Log (R6)")
            print("─" * 40)
            for issue in _report_issues:
                append_issue_log(
                    target_dir=target_dir,
                    source_file=_source_for_logging,
                    output_md=target_md if target_md.exists() else output_md,
                    extractor=_extractor_name,
                    issue_type=issue["issue_type"],
                    severity=issue["severity"],
                    details=issue["details"],
                    action_taken="Continued with warning.",
                )
                print(f"  Logged: [{issue['severity']}] {issue['details'][:60]}")
            print(f"  Issue log: {target_dir / 'CONVERSION-ISSUES.md'}")

        # Also capture cross-validation flags as issues for the report
        if cross_val_flags:
            flag_detail = (f"Cross-validation flagged "
                           f"{len(cross_val_flags)} page(s) with "
                           f">5% word mismatch.")
            _report_issues.append({
                "severity": "WARNING",
                "issue_type": "EXTRACTION_WARNING",
                "details": flag_detail,
            })
            if not args.dry_run:
                append_issue_log(
                    target_dir=target_dir,
                    source_file=_source_for_logging,
                    output_md=target_md if target_md.exists() else output_md,
                    extractor=_extractor_name,
                    issue_type="EXTRACTION_WARNING",
                    severity="WARNING",
                    details=flag_detail,
                    action_taken="Conversion continued. Manual review recommended.",
                )

        # ── R7 + R11: Generate Visual Report ─────────────────────────
        print(f"\n{'─' * 40}")
        print("Step 12: Visual Report (R7)")
        print("─" * 40)

        # Determine overall status
        # _has_frontmatter is now a dedicated boolean from verify_conversion_output()
        # (set at Step 7 above). No need to derive from v_warnings.
        if _report_issues:
            _severities = [i["severity"] for i in _report_issues]
            if "CRITICAL" in _severities:
                _status = "FAILED"
            else:
                _status = "COMPLETE WITH WARNINGS"
        else:
            _status = "COMPLETE"

        report_text = generate_visual_report(
            source_file=args.input_file,
            target_dir=target_dir,
            output_md=target_md if target_md.exists() else output_md,
            source_hash=source_hash,
            extractor=_extractor_name,
            input_stem=input_stem,
            move_actions=_report_move_actions,
            place_actions=_report_place_actions,
            cleanup_actions=cleanup_actions,
            issues=_report_issues,
            verification_passed=verification_passed,
            has_frontmatter=_has_frontmatter,
            dry_run=args.dry_run,
            status=_status,
            image_index_meta=_image_index_meta,
        )

        # Print to stdout
        print(report_text)

        # R6: Write report to disk (R7 spec: always written with --target-dir)
        if not args.dry_run:
            report_filename = (
                f"PIPELINE-REPORT-"
                f"{datetime.now().strftime('%Y%m%d-%H%M%S')}.md"
            )
            report_path = target_dir / report_filename
            try:
                report_path.write_text(report_text, encoding="utf-8")
                print(f"\n  Report written: {report_path}")
            except Exception as e:
                _warn(f"Could not write report file: {e}")
        else:
            print("\n  (DRY RUN) Report not written to disk.")

        # ── Organization Summary ───────────────────────────────────────
        print("\n" + "=" * 60)
        print("ORGANIZATION COMPLETE")
        print("=" * 60)
        print(f"  Target dir:  {target_dir}")
        print(f"  Source:       {originals_dir / args.input_file.name}")
        print(f"  Output:       {target_md}")
        if target_images_dir.exists():
            img_count = sum(1 for _ in target_images_dir.iterdir()
                            if _.is_file())
            print(f"  Images:       {target_images_dir} ({img_count} files)")
        if target_manifest.exists():
            print(f"  Manifest:     {target_manifest}")
        print("=" * 60)

    else:
        # No --target-dir: v3.0 behavior. No organization.
        # R7: Still print a visual report (spec says "always printed").
        # source_hash already computed early (before any file moves)

        _v_passed, _v_warn, _v_err, _v_fm = \
            verify_conversion_output(output_md)

        _ext_name = "unknown"
        if extractor_config:
            _ext_name = extractor_config.extractor
        elif fmt in ("docx", "pptx", "txt"):
            _ext_name = f"convert-office-{fmt}"

        _no_target_status = "COMPLETE"
        if _v_err:
            _no_target_status = "FAILED"
        elif _v_warn:
            _no_target_status = "COMPLETE WITH WARNINGS"

        # R19: Get image_index_meta if it was generated during conversion
        try:
            _img_idx_meta = _image_index_meta
        except NameError:
            _img_idx_meta = None

        report_text = generate_visual_report(
            source_file=args.input_file,
            target_dir=None,
            output_md=output_md,
            source_hash=source_hash,
            extractor=_ext_name,
            input_stem=args.input_file.stem,
            move_actions=[],
            place_actions=[],
            cleanup_actions=[],
            issues=[],
            verification_passed=_v_passed,
            has_frontmatter=_v_fm,
            dry_run=False,
            status=_no_target_status,
            image_index_meta=_img_idx_meta,
        )
        print(report_text)

    # ── R20: Generate Testable Index (when combined with conversion) ──
    # If --generate-testable-index was provided alongside a conversion,
    # run project-level aggregation after conversion completes.
    if args.generate_testable_index is not None:
        project_dir = args.generate_testable_index.resolve()
        if project_dir.exists():
            result = generate_testable_index(project_dir)
            if result:
                print(f"\nTestable index written to: {result}")
            else:
                _warn("Testable index generation produced no output.")
        else:
            _warn(f"Project directory for testable index not found: "
                  f"{project_dir}")


if __name__ == "__main__":
    main()
