#!/usr/bin/env python3
"""convert-paper-marker.py - marker-pdf wrapper for the conversion pipeline.

Thin wrapper that:
1. Calls marker_single CLI to convert PDF to markdown
2. Applies postprocessing (encoding fixes, heading re-leveling, run-togethers)
3. Injects YAML metadata header
4. Outputs pipeline-compatible .md

Does NOT handle image extraction (pipeline handles separately).
"""

import json
import re
import subprocess
import sys
import tempfile
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Optional

try:
    import fitz  # PyMuPDF - for page count and metadata
except ImportError:
    fitz = None

# ─── Constants ────────────────────────────────────────────────────────────────

ENCODING_FIXES = {
    "\u00fe": "+",  # þ (thorn) -> + (font encoding artifact)
    "\u00de": "+",  # Þ (capital thorn) -> +
}

CYRILLIC_CONFUSABLES = {
    "\u0410": "A",  # Cyrillic А -> Latin A
    "\u0412": "B",  # Cyrillic В -> Latin B
    "\u0421": "C",  # Cyrillic С -> Latin C
    "\u0415": "E",  # Cyrillic Е -> Latin E
    "\u041d": "H",  # Cyrillic Н -> Latin H
    "\u041a": "K",  # Cyrillic К -> Latin K
    "\u041c": "M",  # Cyrillic М -> Latin M
    "\u041e": "O",  # Cyrillic О -> Latin O
    "\u0420": "P",  # Cyrillic Р -> Latin P
    "\u0422": "T",  # Cyrillic Т -> Latin T
    "\u0423": "Y",  # Cyrillic У -> Latin Y
    "\u0425": "X",  # Cyrillic Х -> Latin X
}

# Hyphenated compounds common in health economics / HTA literature.
# Each entry is (merged_form_lower, correct_form).
HYPHEN_COMPOUNDS = [
    ("costeffectiveness", "cost-effectiveness"),
    ("costeffective", "cost-effective"),
    ("metaanalysis", "meta-analysis"),
    ("metaanalytic", "meta-analytic"),
    ("metaanalyses", "meta-analyses"),
    ("qualityadjusted", "quality-adjusted"),
    ("progressionfree", "progression-free"),
    ("diseasefree", "disease-free"),
    ("eventfree", "event-free"),
    ("relapsefree", "relapse-free"),
    ("willingnesstopay", "willingness-to-pay"),
    ("intentiontotreat", "intention-to-treat"),
    ("perprotocol", "per-protocol"),
    ("healthrelated", "health-related"),
    ("decisionmaking", "decision-making"),
    ("realworld", "real-world"),
]

# Boilerplate headings to strip entirely (case-insensitive).
BOILERPLATE_HEADINGS = re.compile(
    r"^(ScienceDirect|ORIGINAL\s+RESEARCH|ISPOR\s+Report|Elsevier|"
    r"Taylor\s*&\s*Francis|Springer|Wiley|BMJ|Oxford\s+Academic)$",
    re.IGNORECASE,
)

# Major section patterns that should be H2.
MAJOR_SECTIONS = re.compile(
    r"^(Introduction|Background|Methods|Methodology|Materials?\s+and\s+Methods|"
    r"Results|Discussion|Conclusions?|Summary|References|Acknowledgm?ents?|"
    r"Transparency|Supplementary\s+Materials?|Appendix|ABSTRACT|Keywords?|"
    r"Disclosures?|Funding|Limitations|Conflicts?\s+of\s+Interest|"
    r"Author\s+Contributions?|Data\s+Availability|Ethics)$",
    re.IGNORECASE,
)

ACCENT_MAP = {
    "a": "\u00e1", "e": "\u00e9", "i": "\u00ed",
    "o": "\u00f3", "u": "\u00fa", "n": "\u00f1",
    "A": "\u00c1", "E": "\u00c9", "I": "\u00cd",
    "O": "\u00d3", "U": "\u00da", "N": "\u00d1",
}

# Institutional headers to skip when extracting document title.
# These are generic cover-page headings, not actual document titles.
# Matched case-insensitively as substrings of H1 text.
# SYNC: this list is duplicated in convert-paper.py (INSTITUTIONAL_HEADERS)
#        and run-pipeline.py (_INSTITUTIONAL_HEADERS).
INSTITUTIONAL_HEADERS = [
    "microsoft word", "untitled", "slide ", "powerpoint",
    "health technology assessment",
    "statens legemiddelverk",
    "folkehelseinstituttet",
    "table of contents",
    "erasmus school of health policy",
    "university of oslo",
]
# Exact-match-only headers: these are rejected ONLY when the entire title
# (case-insensitive, stripped) matches exactly. This prevents false positives
# on legitimate titles like "Systematic Review of Markov Models".
INSTITUTIONAL_HEADERS_EXACT = [
    "presentation", "document", "pdf", "contents",
    "systematic review", "rapid review",
    "technology appraisal", "evidence report",
    "clinical practice guideline", "uio",
]

# Journal name patterns for title extraction filtering.
# Prefix match: rejects any heading starting with these strings.
_JOURNAL_PREFIXES = [
    "journal of", "the journal of", "annals of",
    "proceedings of", "transactions on", "archives of",
    "bulletin of", "international journal", "european journal",
    "american journal", "british journal",
]
# Exact match: rejects headings that match these strings exactly.
_JOURNAL_EXACT = {
    "value in health", "pharmacoeconomics",
    "medical decision making", "the lancet", "the bmj",
    "nature", "science", "plos one",
    "bmc health services research", "bmc medicine",
    "health technology assessment",
    "journal of medical economics",
}


# ─── Title/author helpers ────────────────────────────────────────────────────

def _is_journal_name(text: str) -> bool:
    """Check if text is a journal name rather than a paper title.

    Uses prefix matching for common journal name patterns and exact
    matching for specific well-known journal names. Kept separate from
    _is_institutional_header() to avoid false positives on paper titles
    that happen to contain words like 'advances' or 'frontiers'.
    """
    if not text:
        return False
    lower = text.lower().strip()
    for prefix in _JOURNAL_PREFIXES:
        if lower.startswith(prefix):
            return True
    if lower in _JOURNAL_EXACT:
        return True
    return False


def _is_institutional_header(text: str) -> bool:
    """Check if text is a generic institutional header, not a real title.

    Uses case-insensitive substring matching against INSTITUTIONAL_HEADERS
    (genuinely institutional strings like "microsoft word", "untitled").
    Uses exact matching against INSTITUTIONAL_HEADERS_EXACT (generic words
    like "presentation", "document" that appear in real academic titles).
    Also rejects very short texts (<=5 chars) which are usually acronyms
    or page numbers, not titles.
    """
    if not text or len(text) <= 5:
        return True
    text_lower = text.lower().strip()
    # Substring match for genuinely institutional strings
    for header in INSTITUTIONAL_HEADERS:
        if header in text_lower:
            return True
    # Exact match for generic words that appear in real titles
    for header in INSTITUTIONAL_HEADERS_EXACT:
        if text_lower == header:
            return True
    return False


def _clean_heading_text(line: str) -> str:
    """Strip markdown formatting from a heading line, returning plain text."""
    text = re.sub(r"^#+\s*", "", line).strip()
    # Strip bold/italic markers (**, *, ***) from heading text
    text = re.sub(r"\*+", "", text).strip()
    # Strip heading IDs like {#section-id}
    text = re.sub(r"\s*\{#[^}]*\}\s*$", "", text).strip()
    # Strip inline code backticks
    text = text.strip("`").strip()
    return text


def _extract_h1_title(md_text: str) -> Optional[str]:
    """Extract first non-institutional, non-journal H1 heading from markdown.

    Scans the first 30 lines for H1 headings (# Title), skipping any
    that match INSTITUTIONAL_HEADERS or journal name patterns. If all H1
    headings are journal names or institutional headers, falls back to
    the first valid H2 heading. Returns None if no valid heading is found.
    """
    first_h2 = None
    for line in md_text.split("\n")[:30]:
        if line.startswith("# ") and not line.startswith("## "):
            h1_text = _clean_heading_text(line)
            if _is_institutional_header(h1_text):
                continue
            if _is_journal_name(h1_text):
                continue
            return h1_text
        elif line.startswith("## ") and not line.startswith("### ") and first_h2 is None:
            h2_text = _clean_heading_text(line)
            if not _is_institutional_header(h2_text) and not _is_journal_name(h2_text):
                first_h2 = h2_text
    # Fallback: use first valid H2 if no valid H1 found
    return first_h2


# ─── Postprocessing functions ─────────────────────────────────────────────────

def fix_encoding(text: str) -> str:
    """Fix thorn, vulgar-fraction, Cyrillic confusables, accent artifacts."""
    # Thorn/eth -> +
    for old, new in ENCODING_FIXES.items():
        text = text.replace(old, new)

    # Vulgar fraction ¼ -> = (space-guarded to avoid real fractions)
    text = re.sub(r"(?<=\s)\u00bc(?=\s)", "=", text)
    # Also handle ¼ at table cell boundaries (after | or before |)
    text = re.sub(r"(?<=\|)\s*\u00bc\s*(?=\|)", " = ", text)

    # Cyrillic confusables -> Latin
    for cyr, lat in CYRILLIC_CONFUSABLES.items():
        text = text.replace(cyr, lat)

    # Standalone acute accent: e´ -> é etc.
    def _combine_accent(m):
        before = m.group(1)
        after = m.group(2)
        combined = ACCENT_MAP.get(before)
        if combined:
            return combined + after
        return m.group(0)  # no mapping, leave unchanged

    text = re.sub(r"(\w)\u00b4(\w)", _combine_accent, text)

    return text


def fix_ligature_brackets(text: str) -> str:
    """Fix marker ligature-bracket artifacts: 'Stafi[nski' -> 'Stafinski'.

    Marker sometimes splits 'fi'/'fl' ligatures into 'f[i' or 'f[l'.
    The key signal: the bracket is NOT closed (no matching ']') within
    the same word. Real brackets like 'Figure[a]' or 'model[fit]' always
    have a closing bracket, so we only match when the bracket is unclosed
    before whitespace or end-of-string.
    """
    # Match: letter + [ + 1-8 lowercase letters, where no ']' follows
    # before the next whitespace/punctuation/end. This avoids false
    # positives on real bracket notation like Figure[a] or ref[id].
    # {1,8} covers longer words like Koffi[jberg (5), infl[uence (5+).
    text = re.sub(
        r"([A-Za-z])\[([a-z]{1,8})(?=[\s,;:.!?\-/]|$)",
        r"\1\2",
        text,
    )
    return text


def fix_run_togethers(text: str) -> str:
    """Fix known hyphenated compound words merged by OCR/extraction.

    Uses word-boundary anchors to prevent partial matches within longer
    merged strings (e.g., 'realworldevidence' should not become
    'real-worldevidence').
    Preserves original casing pattern: all-upper input stays all-upper,
    title-case input gets title-cased output, lowercase stays lowercase.
    """
    for merged, correct in HYPHEN_COMPOUNDS:
        # Word-boundary anchors prevent partial matches within longer merges
        pattern = re.compile(r"\b" + re.escape(merged) + r"\b", re.IGNORECASE)

        def _case_aware_replace(m, _correct=correct):
            matched = m.group(0)
            if matched.isupper():
                return _correct.upper()
            elif matched[0].isupper():
                return _correct.capitalize()
            return _correct

        text = pattern.sub(_case_aware_replace, text)
    return text


def fix_headings(text: str) -> str:
    """Re-level headings to consistent H1 -> H2 -> H3 -> H4 hierarchy."""
    lines = text.split("\n")
    heading_re = re.compile(r"^(#{1,6})\s+(.+)$")

    # Pass 1: parse all headings
    headings = []  # (line_idx, level, text_stripped)
    for i, line in enumerate(lines):
        m = heading_re.match(line)
        if m:
            headings.append((i, len(m.group(1)), m.group(2).strip()))

    # Helper: strip bold/italic markers for pattern matching only.
    # Headings like '### **Introduction**' need the ** stripped before
    # matching against MAJOR_SECTIONS / BOILERPLATE_HEADINGS.
    def _strip_markers(txt: str) -> str:
        return txt.strip("* ").strip()

    if not headings:
        return text

    # Pass 2: classify each heading
    title_idx = None
    for idx, (line_i, level, htxt) in enumerate(headings):
        # Strip bold/italic markers for pattern matching
        htxt_clean = _strip_markers(htxt)
        # Skip boilerplate
        if BOILERPLATE_HEADINGS.match(htxt_clean):
            lines[line_i] = ""  # remove boilerplate heading
            headings[idx] = (line_i, 0, htxt)  # mark as removed
            continue
        # First non-boilerplate heading with >15 chars is likely the title
        if title_idx is None and len(htxt_clean) > 15:
            title_idx = idx
            continue
        # If still no title and we hit a major section, title was short
        if title_idx is None and MAJOR_SECTIONS.match(htxt_clean):
            # Previous heading (if any non-removed) is probably title
            for prev in range(idx - 1, -1, -1):
                if headings[prev][1] != 0:
                    title_idx = prev
                    break
            if title_idx is None:
                title_idx = idx  # fallback: this IS the title

    # Pass 3: assign target levels
    # title -> H1, major sections -> H2, everything else relative to context
    #
    # We track the last major-section boundary to determine relative depth.
    # After a major section (H2), the first sub-heading goes to H3,
    # subsequent deeper headings go to H4. Original relative ordering is
    # preserved: if orig_level increases, target increases (up to H4).
    last_major_level = 2  # after title, next context is major-section depth
    last_orig_level = None  # track original level for relative depth
    last_target = 1
    for idx, (line_i, orig_level, htxt) in enumerate(headings):
        if orig_level == 0:
            continue  # removed boilerplate

        # Strip bold/italic markers for pattern matching
        htxt_clean = _strip_markers(htxt)
        if idx == title_idx:
            target = 1
        elif MAJOR_SECTIONS.match(htxt_clean):
            target = 2
            last_major_level = 2
            last_orig_level = orig_level
        elif orig_level == 1:
            # Non-title H1 that is not a major section: demote to H2
            target = 2
            last_major_level = 2
            last_orig_level = orig_level
        else:
            # Sub-heading: assign H3 or H4 based on depth relative to
            # the last major section. If this heading was deeper than
            # the previous heading in the original, go one level deeper.
            if last_orig_level is not None and orig_level > last_orig_level:
                target = min(last_target + 1, 4)
            else:
                target = last_major_level + 1  # typically H3
            target = min(target, 4)

        target = min(target, 4)  # never go beyond H4
        lines[line_i] = "#" * target + " " + htxt
        last_target = target
        if orig_level != 0:
            last_orig_level = orig_level

    return "\n".join(lines)


def clean_html_spans(text: str) -> str:
    """Remove marker's HTML span tags (page markers, etc.)."""
    text = re.sub(r'<span\s+id="[^"]*"\s*>\s*</span>\s*', "", text)
    return text


def fix_references(text: str) -> str:
    """Fix '- 0[N]' prefix artifact in reference lists."""
    text = re.sub(r"^- 0\[", "- [", text, flags=re.MULTILINE)
    return text


# ─── Marker invocation ────────────────────────────────────────────────────────

def run_marker(input_pdf: Path, temp_dir: Path,
               timeout_per_attempt: int = 600) -> Path:
    """Call marker_single CLI. Returns path to output .md file.

    If the first attempt fails (e.g. MPS crash on Apple Silicon), retries
    once with TORCH_DEVICE=cpu to bypass GPU-related upstream bugs.

    Args:
        input_pdf: Path to the source PDF.
        temp_dir: Temporary directory for marker output.
        timeout_per_attempt: Timeout in seconds per attempt (default 600).
                             Each attempt (MPS + CPU retry) uses this value.
    """
    import os

    cmd = [
        "marker_single", str(input_pdf),
        "--output_dir", str(temp_dir),
        "--output_format", "markdown",
        "--disable_image_extraction",
        "--disable_tqdm",
    ]
    try:
        result = subprocess.run(cmd, capture_output=True, text=True,
                                timeout=timeout_per_attempt)
    except subprocess.TimeoutExpired:
        print(
            f"WARNING: marker_single timed out after "
            f"{timeout_per_attempt}s (MPS attempt)",
            file=sys.stderr,
        )
        # Fall through to CPU retry below
        result = None
    except FileNotFoundError:
        print("marker_single not found on PATH", file=sys.stderr)
        sys.exit(3)  # exit 3 = extractor failure (pipeline can fallback)
    if result is None or result.returncode != 0:
        # MPS/GPU crashed or timed out - retry with CPU
        exit_info = f"exit {result.returncode}" if result else "timeout"
        print(
            f"marker_single failed ({exit_info}), "
            f"retrying with TORCH_DEVICE=cpu...",
            file=sys.stderr,
        )
        cpu_env = os.environ.copy()
        cpu_env["TORCH_DEVICE"] = "cpu"
        try:
            result = subprocess.run(
                cmd, capture_output=True, text=True, env=cpu_env,
                timeout=timeout_per_attempt,
            )
        except subprocess.TimeoutExpired:
            print(
                f"WARNING: marker_single timed out after "
                f"{timeout_per_attempt}s on CPU fallback",
                file=sys.stderr,
            )
            sys.exit(3)  # exit 3 = extractor failure (pipeline can fallback)
        if result.returncode != 0:
            print(
                f"marker_single failed on CPU fallback too "
                f"(exit {result.returncode}):",
                file=sys.stderr,
            )
            print(result.stderr, file=sys.stderr)
            sys.exit(3)  # exit 3 = extractor failure (pipeline can fallback)
        print("CPU fallback succeeded.", file=sys.stderr)

    # Find the .md output (handles nesting: Stem/Stem.md)
    md_files = list(temp_dir.rglob("*.md"))
    if not md_files:
        print("No .md output found from marker_single", file=sys.stderr)
        sys.exit(3)  # exit 3 = extractor failure
    return md_files[0]


def read_marker_meta(md_path: Path) -> dict:
    """Read marker's meta.json if present (sibling of .md file)."""
    meta_path = md_path.parent / (md_path.stem + "_meta.json")
    if meta_path.exists():
        return json.loads(meta_path.read_text(encoding="utf-8"))
    return {}


# ─── YAML header ──────────────────────────────────────────────────────────────

def build_yaml_header(input_pdf: Path, pdf_meta: dict) -> str:
    """Build YAML frontmatter compatible with existing pipeline.

    pdf_meta dict must contain page_count (int), and optionally title and
    author (str or None). Title and author are included only when present.
    """
    header = "---\n"
    header += f"source_file: {input_pdf.name}\n"
    header += f'source_path: "{input_pdf.resolve()}"\n'
    header += "source_format: pdf\n"
    if pdf_meta.get("title"):
        safe_title = pdf_meta["title"].replace("\\", "\\\\").replace('"', '\\"')
        header += f'title: "{safe_title}"\n'
    if pdf_meta.get("author"):
        safe_author = pdf_meta["author"].replace("\\", "\\\\").replace('"', '\\"')
        header += f'author: "{safe_author}"\n'
    header += f"pages: {pdf_meta['page_count']}\n"
    header += f"conversion_date: {datetime.now().strftime('%Y-%m-%d')}\n"
    header += "conversion_tool: marker-pdf + PyMuPDF\n"
    header += "fidelity_standard: verbatim (QC required)\n"
    header += "document_type: research_paper\n"
    header += "image_notes: pending\n"
    header += "---\n\n"
    return header


def get_pdf_metadata(input_pdf: Path, meta: dict) -> dict:
    """Get page count, title, and author from fitz (preferred) or marker meta.

    Returns dict with keys: page_count (int), title (str|None), author (str|None).
    Title from fitz is validated: must be >5 chars and not an institutional header.
    Author from fitz is included as-is if >1 char.
    """
    result = {"page_count": 0, "title": None, "author": None}
    if fitz:
        doc = None
        try:
            doc = fitz.open(str(input_pdf))
            result["page_count"] = len(doc)
            pdf_meta = doc.metadata or {}
            # Title: from PDF properties, validated
            _title = (pdf_meta.get("title") or "").strip()
            if (_title and len(_title) > 5
                    and not _is_institutional_header(_title)
                    and not _is_journal_name(_title)):
                result["title"] = _title
            # Author: from PDF properties
            _author = (pdf_meta.get("author") or "").strip()
            if _author and len(_author) > 1:
                result["author"] = _author
            return result
        except Exception:
            pass
        finally:
            if doc is not None:
                doc.close()
    # Fallback: marker meta for page count (no title/author available)
    page_stats = meta.get("page_stats", [])
    if page_stats:
        result["page_count"] = len(page_stats)
    return result


# ─── Main ─────────────────────────────────────────────────────────────────────

def convert(input_pdf: Path, output_dir: Path) -> Path:
    """Full conversion pipeline: marker -> postprocess -> YAML -> write."""
    stem = input_pdf.stem
    output_dir.mkdir(parents=True, exist_ok=True)

    # 1. Run marker to temp directory
    # Scale timeout for large PDFs: base 600s + 3s per page above 100
    _marker_timeout = 600
    try:
        import fitz as _fitz_timeout
        _doc_timeout = _fitz_timeout.open(str(input_pdf))
        _page_count = len(_doc_timeout)
        _doc_timeout.close()
        if _page_count > 100:
            _marker_timeout = 600 + (_page_count - 100) * 3
            print(f"Large PDF ({_page_count} pages): "
                  f"timeout scaled to {_marker_timeout}s per attempt")
    except Exception:
        pass  # fitz unavailable or PDF unreadable; use default

    with tempfile.TemporaryDirectory(prefix="marker_") as tmp:
        tmp_path = Path(tmp)
        print(f"Running marker_single on {input_pdf.name}...")
        md_path = run_marker(input_pdf, tmp_path,
                             timeout_per_attempt=_marker_timeout)
        md_text = md_path.read_text(encoding="utf-8")
        meta = read_marker_meta(md_path)

    # 2. Get PDF metadata (page count, title, author from fitz)
    pdf_meta = get_pdf_metadata(input_pdf, meta)

    # 3. Postprocess
    md_text = unicodedata.normalize("NFC", md_text)
    md_text = fix_encoding(md_text)
    md_text = fix_ligature_brackets(md_text)
    md_text = fix_run_togethers(md_text)
    md_text = clean_html_spans(md_text)
    md_text = fix_references(md_text)
    md_text = fix_headings(md_text)

    # 4. Enrich title from markdown H1 (more reliable than fitz metadata)
    h1_title = _extract_h1_title(md_text)
    if h1_title:
        pdf_meta["title"] = h1_title

    # 5. Prepend YAML header
    md_text = build_yaml_header(input_pdf, pdf_meta) + md_text

    # 6. Write output
    out_path = output_dir / f"{stem}.md"
    out_path.write_text(md_text, encoding="utf-8")
    print(f"Written: {out_path} ({pdf_meta['page_count']} pages)")
    return out_path


def main():
    import argparse

    parser = argparse.ArgumentParser(
        description="Convert PDF to markdown using marker-pdf with postprocessing.",
    )
    parser.add_argument("input_pdf", type=Path, help="Path to input PDF file")
    parser.add_argument(
        "--output-dir", type=Path, default=None,
        help="Output directory (default: same as input PDF)",
    )
    args = parser.parse_args()

    if not args.input_pdf.exists():
        print(f"File not found: {args.input_pdf}", file=sys.stderr)
        sys.exit(1)

    output_dir = args.output_dir or args.input_pdf.parent
    convert(args.input_pdf, output_dir)


if __name__ == "__main__":
    main()
