#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Convert documents (PDF, DOCX, PPTX, XLSX) to high-fidelity Markdown with enhanced metadata.

MODIFIED from ~/.claude/scripts/convert-paper.py with 6 NEW capabilities:
1. Context summary generation (extractive, no LLM)
2. Section-to-image mapping in manifest
3. Caption detection for PDF images
4. Document domain auto-detection
5. Duplicate xref tracking (bug fix)
6. Enhanced manifest fields (type_guess, section_context, detected_caption)

PRESERVED capabilities:
- Multi-panel image splitting (>800x800)
- Sparse page rendering (300 DPI for pages with <200 chars text)
- pymupdf.layout mode import
- MarkItDown PDF fallback

Text extraction:
  - PDF: pymupdf4llm (proper tables, headings, multi-column support)
  - DOCX/PPTX/XLSX: MarkItDown (good for structured formats)
Image extraction: PyMuPDF (PDF), python-docx (DOCX), python-pptx (PPTX).

Runs ENTIRELY via Bash/Python, bypassing Claude's API copyright filter.

Usage:
    python3 convert-paper.py <input_file> [-o output.md] [-i images/] [-s short-name]

Supported: .pdf .docx .pptx .xlsx
"""

import argparse
import io
import json
import os
import re
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional

try:
    from markitdown import MarkItDown
except ImportError:
    MarkItDown = None

try:
    import pymupdf4llm
    # Layout mode (pymupdf.layout) is intentionally NOT imported.
    # It activates AI-based page analysis for better headings,
    # multi-column, and table detection — but it also triggers a
    # known pymupdf4llm bug on certain PDFs where table bounding
    # boxes are empty, causing:
    #   ValueError: min() iterable argument is empty
    #   (at pymupdf/table.py:1534)
    # This crash was observed on a large PDF with 166 images.
    # Without layout mode, to_markdown() still works
    # correctly using standard text extraction.  The ValueError
    # catch at the to_markdown() call site (exit code 3) remains
    # as a safety net in case other pymupdf4llm bugs surface.
    _HAS_LAYOUT = False
except ImportError:
    pymupdf4llm = None
    _HAS_LAYOUT = False

try:
    from PIL import Image
except ImportError:
    Image = None
    print("WARNING: Pillow not installed. Multi-panel splitting disabled.")

SUPPORTED_FORMATS = {".pdf", ".docx", ".pptx", ".xlsx"}
DPI = 300
ZOOM = DPI / 72

# Institutional headers to skip when extracting document title.
# These are generic cover-page headings, not actual document titles.
# Matched case-insensitively as substrings of H1 text.
# SYNC: this list is duplicated in run-pipeline.py (_INSTITUTIONAL_HEADERS)
INSTITUTIONAL_HEADERS = [
    "health technology assessment",
    "statens legemiddelverk",
    "folkehelseinstituttet",
    "table of contents",
    "contents",
    "systematic review",
    "rapid review",
    "technology appraisal",
    "evidence report",
    "clinical practice guideline",
]

# Domain detection keywords for auto-classification
DOMAIN_KEYWORDS = {
    "health_economics": [
        "cost", "QALY", "ICER", "Markov", "willingness-to-pay",
        "cost-effectiveness", "incremental", "budget impact", "threshold"
    ],
    "clinical_trial": [
        "randomized", "placebo", "endpoint", "ITT", "CONSORT",
        "adverse event", "RCT", "trial", "phase", "randomization"
    ],
    "systematic_review": [
        "PRISMA", "meta-analysis", "forest plot", "I-squared",
        "pooled", "heterogeneity", "systematic", "review"
    ],
    "hta_regulatory": [
        "NICE", "DMP", "NoMA", "reimbursement", "submission",
        "appraisal", "HTA", "regulatory", "guideline"
    ],
    "epidemiology": [
        "incidence", "prevalence", "registry", "cohort", "DALY",
        "mortality", "population", "burden", "disease"
    ],
    "methodology": [
        "model validation", "simulation", "calibration",
        "structural uncertainty", "sensitivity", "probabilistic"
    ],
}

# High-specificity override keywords that force a domain match.
# Checked BEFORE frequency scoring. If any keyword is found (case-insensitive
# substring match), that domain wins immediately. These are terms that are
# unambiguous domain signals — e.g. "NoMA" only appears in HTA regulatory.
# Note: "ICER" (bare) is in health_economics; "ICER threshold" is in
# hta_regulatory (the phrase "ICER threshold" implies a policy/appraisal context).
DOMAIN_OVERRIDE_KEYWORDS = {
    "hta_regulatory": [
        "NoMA", "metodevurdering", "ICER threshold", "cost per QALY",
        "NICE appraisal", "health technology assessment",
        "reimbursement decision", "HTA body", "DMP",
    ],
    "health_economics": [
        "Markov model", "cost-effectiveness analysis",
        "willingness to pay", "ICER",
        "incremental cost-effectiveness",
        "cost-utility analysis",
        "survival analysis",
        "hazard ratio",
    ],
}


# ═══════════════════════════════════════════════════════════════════════════
# FORMAT DETECTION
# ═══════════════════════════════════════════════════════════════════════════

def detect_format(file_path: Path) -> str:
    """Return normalized format string: pdf, docx, pptx, xlsx."""
    ext = file_path.suffix.lower()
    if ext not in SUPPORTED_FORMATS:
        print(f"ERROR: Unsupported format '{ext}'. Supported: {SUPPORTED_FORMATS}")
        sys.exit(1)
    return ext.lstrip(".")


def get_page_count(file_path: Path, fmt: str) -> Optional[int]:
    """Get page/slide/sheet count where possible."""
    if fmt == "pdf":
        import fitz
        doc = fitz.open(str(file_path))
        count = len(doc)
        doc.close()
        return count
    elif fmt == "pptx":
        from pptx import Presentation
        prs = Presentation(str(file_path))
        return len(prs.slides)
    elif fmt == "xlsx":
        import openpyxl
        try:
            wb = openpyxl.load_workbook(str(file_path), read_only=True)
            count = len(wb.sheetnames)
            wb.close()
            return count
        except Exception:
            return None
    elif fmt == "docx":
        # DOCX doesn't have a reliable page count without rendering
        return None
    return None


# ═══════════════════════════════════════════════════════════════════════════
# IMAGE EXTRACTION - PDF
# ═══════════════════════════════════════════════════════════════════════════

def detect_caption_near_image(page, img_rect) -> Optional[str]:
    """
    Detect "Figure N:" or "Table N:" text near the image bounding box.
    Searches within 100 pixels below the image.
    """
    import fitz
    try:
        # Expand search area below image
        search_rect = fitz.Rect(
            img_rect.x0 - 50,
            img_rect.y1,
            img_rect.x1 + 50,
            img_rect.y1 + 100
        )
        text = page.get_text("text", clip=search_rect)
        # Look for common caption patterns
        match = re.search(r'(Figure|Table|Fig\.|Tab\.)\s*\d+[:\.]?\s*([^\n]{0,80})', text, re.IGNORECASE)
        if match:
            return match.group(0).strip()
    except Exception:
        pass
    return None


def guess_type_from_caption(caption: Optional[str]) -> Optional[str]:
    """
    Guess image type from detected caption text.
    Returns a type string compatible with IMAGE NOTE types.
    """
    if not caption:
        return None

    caption_lower = caption.lower()

    # Map common caption keywords to image types
    type_map = {
        "forest plot": "forest-plot",
        "kaplan": "kaplan-meier",
        "survival": "kaplan-meier",
        "tornado": "tornado-diagram",
        "decision tree": "decision-tree",
        "flow chart": "flow-chart",
        "flowchart": "flow-chart",
        "scatter": "scatter",
        "cost-effectiveness plane": "scatter",
        "box plot": "box-plot",
        "histogram": "histogram",
        "heatmap": "heatmap",
        "bar chart": "bar-chart",
        "line chart": "line-chart",
        "pie chart": "pie-chart",
        "funnel": "funnel-plot",
        "network": "network-diagram",
    }

    for keyword, img_type in type_map.items():
        if keyword in caption_lower:
            return img_type

    # Default: table vs figure
    if "table" in caption_lower:
        return "table-image"
    elif "figure" in caption_lower or "fig." in caption_lower:
        return "other"  # Unknown figure type

    return None


def extract_images_pdf(file_path: Path, images_dir: Path,
                       short_name: str, md_text: str) -> list[dict]:
    """Extract images from PDF via PyMuPDF with enhanced metadata."""
    import fitz

    doc = fitz.open(str(file_path))
    images_dir.mkdir(parents=True, exist_ok=True)
    image_index = []
    figure_num = 0
    extracted_xrefs = set()  # Track extracted xrefs to avoid duplicates

    # Parse section structure from md_text for section context
    sections = parse_section_structure(md_text)

    for page_idx in range(len(doc)):
        page = doc[page_idx]
        page_images = page.get_images(full=True)

        if page_images:
            for img_info in page_images:
                xref = img_info[0]

                # Skip if already extracted (duplicate detection fix)
                if xref in extracted_xrefs:
                    continue

                try:
                    base_image = doc.extract_image(xref)
                except Exception:
                    continue
                if base_image is None:
                    continue

                img_bytes = base_image["image"]
                img_ext = base_image.get("ext", "png")
                width = base_image.get("width", 0)
                height = base_image.get("height", 0)

                if width < 50 or height < 50:
                    continue

                extracted_xrefs.add(xref)  # Mark as extracted
                figure_num += 1
                filename = f"{short_name}-fig{figure_num}-page{page_idx + 1}.{img_ext}"
                filepath = images_dir / filename

                with open(filepath, "wb") as f:
                    f.write(img_bytes)

                # ── Blank detection ───────────────────────────────────────────────
                # Detect near-black fragments, pure-color blocks, and other blank
                # placeholder images that waste Opus vision tokens.
                # Aligned with run-pipeline.py _is_blank_image() canonical logic:
                # Tier 1: file size < 2KB (tiny placeholder)
                # Tier 2: file_size < 5KB AND mean < 30 AND unique_colors < 16
                #         (near-black with few colors — OR tier, not AND)
                # Tier 3: std < 5.0 (near-uniform pixel values)
                # M1: unique_colors < 32 AND mean < 240 AND std < 30
                #     (color-block / gradient bar detection)
                # Applied at write time so the manifest field is populated before
                # prepare-image-analysis.py reads it.
                _is_blank = False
                _blank_reason = None
                try:
                    _file_size = os.path.getsize(str(filepath))
                    # Tier 1: tiny placeholder
                    if _file_size < 2000:
                        _is_blank = True
                        _blank_reason = f"file_size={_file_size}B"
                    elif Image is not None:
                        import numpy as _np
                        _pil_img = Image.open(str(filepath))
                        _gray = _pil_img.convert("L")
                        _arr = _np.array(_gray)
                        _std = float(_arr.std())
                        _mean = float(_arr.mean())

                        # Compute unique colors for Tier 2 and M1
                        _unique_colors = -1
                        if _pil_img.mode in ('RGB', 'RGBA'):
                            _rgb = _pil_img.convert('RGB')
                            _w_px, _h_px = _rgb.size
                            if _w_px * _h_px > 50000:
                                _rgb = _rgb.resize(
                                    (min(_w_px, 250), min(_h_px, 200)),
                                    Image.NEAREST)
                            _unique_colors = len(set(_rgb.getdata()))

                        _pil_img.close()

                        # Tier 2: near-black with few colors (OR tier)
                        if (_file_size > 0 and _file_size < 5000
                                and _mean < 30
                                and _unique_colors >= 0
                                and _unique_colors < 16):
                            _is_blank = True
                            _blank_reason = (
                                f"near_black(mean={_mean:.2f},"
                                f"colors={_unique_colors},"
                                f"size={_file_size}B)")
                        # Tier 3: near-uniform pixel values
                        elif _std < 5.0:
                            _is_blank = True
                            _blank_reason = f"std={_std:.2f}"
                        # M1: color-block / gradient bar detection
                        elif (_unique_colors >= 0
                                and _unique_colors < 32
                                and _mean < 240
                                and _std < 30):
                            _is_blank = True
                            _blank_reason = (
                                f"color_block(colors={_unique_colors},"
                                f"mean={_mean:.2f},std={_std:.2f})")
                except Exception:
                    pass  # If detection fails, leave _is_blank=False (safe default)
                # ─────────────────────────────────────────────────────────────────

                # Get image rect for caption detection
                img_rects = page.get_image_rects(xref)
                img_rect = img_rects[0] if img_rects else None
                detected_caption = detect_caption_near_image(page, img_rect) if img_rect else None
                type_guess = guess_type_from_caption(detected_caption)

                # Get nearby text context (200 chars before and after image position)
                nearby_text = get_text_near_image(page, img_rect) if img_rect else None

                # Find section context
                section_context = find_section_for_page(sections, page_idx + 1, total_pages=len(doc))

                image_index.append({
                    "page": page_idx + 1,
                    "figure_num": figure_num,
                    "filename": filename,
                    "description": f"Figure {figure_num} from page {page_idx + 1}",
                    "width": width,
                    "height": height,
                    "source_format": "pdf_embedded",
                    "detected_caption": detected_caption,
                    "type_guess": type_guess,
                    "section_context": section_context,
                    "nearby_text": nearby_text,
                    "analysis_status": "pending",
                    "blank": _is_blank,
                    "blank_reason": _blank_reason,
                    "decorative": False,
                    "is_duplicate": False,
                })

                # Split multi-panel if large enough (PRESERVED capability)
                if Image is not None and (width > 800 and height > 800):
                    try:
                        _split_panels(img_bytes, figure_num, page_idx + 1,
                                      short_name, images_dir, image_index)
                    except Exception:
                        pass

    # Render sparse/figure-only pages (PRESERVED capability)
    pages_with_images = {e["page"] for e in image_index}
    for page_idx in range(len(doc)):
        page_num = page_idx + 1
        if page_num in pages_with_images:
            continue
        page = doc[page_idx]
        text = page.get_text("text").strip()
        if len(text) > 200:
            continue

        figure_num += 1
        filename = f"{short_name}-page{page_num}-render-300dpi.png"
        filepath = images_dir / filename
        mat = fitz.Matrix(ZOOM, ZOOM)
        pix = page.get_pixmap(matrix=mat)
        pix.save(str(filepath))

        section_context = find_section_for_page(sections, page_num, total_pages=len(doc))

        image_index.append({
            "page": page_num,
            "figure_num": figure_num,
            "filename": filename,
            "description": f"Full page render of page {page_num} at 300 DPI",
            "width": pix.width,
            "height": pix.height,
            "source_format": "pdf_page_render",
            "detected_caption": None,
            "type_guess": "other",
            "section_context": section_context,
            "nearby_text": None,
            "analysis_status": "pending",
        })

    doc.close()
    return image_index


def get_text_near_image(page, img_rect) -> Optional[str]:
    """Extract text surrounding the image (200 chars before and after)."""
    import fitz
    if not img_rect:
        return None
    try:
        # Get all text on page
        full_text = page.get_text("text")

        # Get text in expanded rect around image
        expanded_rect = fitz.Rect(
            max(0, img_rect.x0 - 100),
            max(0, img_rect.y0 - 100),
            img_rect.x1 + 100,
            img_rect.y1 + 100
        )
        nearby = page.get_text("text", clip=expanded_rect)

        # Clean and truncate
        nearby = re.sub(r'\s+', ' ', nearby).strip()
        if len(nearby) > 200:
            nearby = nearby[:200] + "..."

        return nearby if nearby else None
    except Exception:
        return None


def find_section_for_page(sections: list[dict], page_num: int,
                          total_pages: int = 0) -> dict:
    """
    Find the section that contains the given page number.

    Uses a document-calibrated lines-per-page ratio (total_lines / total_pages)
    instead of a hardcoded constant, then maps page_num to an estimated line
    number and finds which section contains that line.

    Args:
        sections: Parsed section structure from parse_section_structure().
        page_num: 1-based page number in the PDF.
        total_pages: Total number of pages in the PDF (from len(doc)).
            When provided, enables accurate per-document calibration.
            Falls back to 45 lines/page when 0 or not provided.
    """
    if not sections:
        return {
            "heading": "Unknown Section",
            "heading_level": 2,
        }

    # Estimate total lines from all sections
    total_lines = max((s.get("line_end") or 0) for s in sections)
    if total_lines == 0:
        # Fallback to first section
        return {
            "heading": sections[0]["heading"],
            "heading_level": sections[0]["level"],
        }

    # Use document-calibrated ratio instead of constant 45 lines/page.
    # The old constant caused all images on long documents (200+ pages)
    # to overshoot section boundaries and fall back to the last heading.
    if total_pages > 0:
        estimated_lines_per_page = total_lines / total_pages
    else:
        estimated_lines_per_page = 45  # legacy fallback
    estimated_line_num = page_num * estimated_lines_per_page

    # Find section containing this estimated line
    for section in sections:
        line_start = section.get("line_start", 0)
        line_end = section.get("line_end", 0)
        if line_start <= estimated_line_num <= line_end:
            return {
                "heading": section["heading"],
                "heading_level": section["level"],
            }

    # If no match, find the nearest preceding section (closest line_end
    # that is <= estimated_line_num). This is more accurate than always
    # returning the last section.
    best_section = sections[-1]  # default fallback
    best_distance = float('inf')
    for section in sections:
        line_end = section.get("line_end", 0)
        if line_end <= estimated_line_num:
            distance = estimated_line_num - line_end
            if distance < best_distance:
                best_distance = distance
                best_section = section

    return {
        "heading": best_section["heading"],
        "heading_level": best_section["level"],
    }


# ═══════════════════════════════════════════════════════════════════════════
# IMAGE EXTRACTION - DOCX
# ═══════════════════════════════════════════════════════════════════════════

def extract_images_docx(file_path: Path, images_dir: Path,
                        short_name: str) -> list[dict]:
    """Extract embedded images from DOCX via python-docx."""
    from docx import Document

    doc = Document(str(file_path))
    images_dir.mkdir(parents=True, exist_ok=True)
    image_index = []
    figure_num = 0

    for rel in doc.part.rels.values():
        if "image" in rel.reltype:
            try:
                img_data = rel.target_part.blob
                content_type = rel.target_part.content_type
            except Exception:
                continue

            # Determine extension
            ext_map = {
                "image/png": "png",
                "image/jpeg": "jpg",
                "image/gif": "gif",
                "image/bmp": "bmp",
                "image/tiff": "tiff",
                "image/x-emf": "emf",
                "image/x-wmf": "wmf",
            }
            img_ext = ext_map.get(content_type, "png")

            # Skip vector formats that can't be displayed as raster
            if img_ext in ("emf", "wmf"):
                continue

            # Get dimensions if possible
            width, height = 0, 0
            if Image is not None and img_ext in ("png", "jpg", "gif", "bmp"):
                try:
                    img = Image.open(io.BytesIO(img_data))
                    width, height = img.size
                except Exception:
                    pass

            # Skip tiny images
            if width > 0 and width < 50 and height < 50:
                continue

            figure_num += 1
            filename = f"{short_name}-fig{figure_num}.{img_ext}"
            filepath = images_dir / filename

            with open(filepath, "wb") as f:
                f.write(img_data)

            image_index.append({
                "page": None,  # DOCX doesn't have reliable page numbers
                "figure_num": figure_num,
                "filename": filename,
                "description": f"Figure {figure_num} from document",
                "width": width,
                "height": height,
                "source_format": "docx_embedded",
                "detected_caption": None,
                "type_guess": None,
                "section_context": None,
                "nearby_text": None,
                "analysis_status": "pending",
            })

    return image_index


# ═══════════════════════════════════════════════════════════════════════════
# IMAGE EXTRACTION - PPTX
# ═══════════════════════════════════════════════════════════════════════════

def extract_images_pptx(file_path: Path, images_dir: Path,
                        short_name: str) -> list[dict]:
    """
    Extract from PPTX:
    1. All embedded images from shapes
    2. Full slide renders at 300 DPI (via PyMuPDF PDF conversion)
    """
    from pptx import Presentation

    prs = Presentation(str(file_path))
    images_dir.mkdir(parents=True, exist_ok=True)
    image_index = []
    figure_num = 0

    # Extract embedded images from shapes
    for slide_idx, slide in enumerate(prs.slides):
        slide_num = slide_idx + 1
        for shape in slide.shapes:
            if shape.shape_type == 13:  # Picture
                try:
                    img_data = shape.image.blob
                    content_type = shape.image.content_type
                except Exception:
                    continue

                ext_map = {
                    "image/png": "png",
                    "image/jpeg": "jpg",
                    "image/gif": "gif",
                    "image/bmp": "bmp",
                }
                img_ext = ext_map.get(content_type, "png")

                width, height = 0, 0
                if Image is not None:
                    try:
                        img = Image.open(io.BytesIO(img_data))
                        width, height = img.size
                    except Exception:
                        pass

                if width > 0 and width < 30 and height < 30:
                    continue

                figure_num += 1
                filename = f"{short_name}-slide{slide_num}-fig{figure_num}.{img_ext}"
                filepath = images_dir / filename

                with open(filepath, "wb") as f:
                    f.write(img_data)

                image_index.append({
                    "page": slide_num,
                    "figure_num": figure_num,
                    "filename": filename,
                    "description": f"Figure {figure_num} from slide {slide_num}",
                    "width": width,
                    "height": height,
                    "source_format": "pptx_embedded",
                    "detected_caption": None,
                    "type_guess": None,
                    "section_context": None,
                    "nearby_text": None,
                    "analysis_status": "pending",
                })

    # Render full slides via libreoffice -> PDF -> PyMuPDF (if available)
    # (PRESERVED capability)
    try:
        import fitz
        import subprocess
        import tempfile

        with tempfile.TemporaryDirectory() as tmpdir:
            # Convert PPTX to PDF via libreoffice
            result = subprocess.run(
                ["libreoffice", "--headless", "--convert-to", "pdf",
                 "--outdir", tmpdir, str(file_path)],
                capture_output=True, timeout=120
            )
            if result.returncode == 0:
                pdf_files = list(Path(tmpdir).glob("*.pdf"))
                if pdf_files:
                    doc = fitz.open(str(pdf_files[0]))
                    for page_idx in range(len(doc)):
                        slide_num = page_idx + 1
                        figure_num += 1
                        filename = f"{short_name}-slide{slide_num}-render-300dpi.png"
                        filepath = images_dir / filename
                        mat = fitz.Matrix(ZOOM, ZOOM)
                        pix = doc[page_idx].get_pixmap(matrix=mat)
                        pix.save(str(filepath))
                        image_index.append({
                            "page": slide_num,
                            "figure_num": figure_num,
                            "filename": filename,
                            "description": f"Full slide {slide_num} render at 300 DPI",
                            "width": pix.width,
                            "height": pix.height,
                            "source_format": "pptx_slide_render",
                            "detected_caption": None,
                            "type_guess": None,
                            "section_context": None,
                            "nearby_text": None,
                            "analysis_status": "pending",
                        })
                    doc.close()
    except (ImportError, FileNotFoundError, subprocess.TimeoutExpired):
        print("  WARNING: libreoffice not found. Skipping slide renders.")
        print("  (Embedded images still extracted)")

    return image_index


# ═══════════════════════════════════════════════════════════════════════════
# MULTI-PANEL SPLITTING (PRESERVED)
# ═══════════════════════════════════════════════════════════════════════════

def _split_panels(img_bytes: bytes, figure_num: int, page_num: int,
                  short_name: str, images_dir: Path,
                  image_index: list[dict]):
    """Split a multi-panel figure (A/B) into individual panels."""
    img = Image.open(io.BytesIO(img_bytes))
    w, h = img.size

    if w > h:
        panel_a = img.crop((0, 0, w // 2, h))
        panel_b = img.crop((w // 2, 0, w, h))
    else:
        panel_a = img.crop((0, 0, w, h // 2))
        panel_b = img.crop((0, h // 2, w, h))

    for label, panel in [("a", panel_a), ("b", panel_b)]:
        filename = f"{short_name}-fig{figure_num}{label}-page{page_num}.png"
        filepath = images_dir / filename
        panel.save(str(filepath))
        image_index.append({
            "page": page_num,
            "figure_num": f"{figure_num}{label}",
            "filename": filename,
            "description": f"Figure {figure_num} panel {label.upper()} from page {page_num}",
            "width": panel.width,
            "height": panel.height,
            "source_format": "panel_split",
            "detected_caption": None,
            "type_guess": None,
            "section_context": None,
            "nearby_text": None,
            "analysis_status": "pending",
        })


# ═══════════════════════════════════════════════════════════════════════════
# TEXT EXTRACTION + ASSEMBLY
# ═══════════════════════════════════════════════════════════════════════════

def extract_text(file_path: Path, fmt: str,
                 extractor: str = "pymupdf4llm") -> str:
    """Extract text from document.

    PDF: pymupdf4llm (proper tables, headings, multi-column).
    PDF+tesseract: Tesseract OCR via PyMuPDF OCR bridge (scanned PDFs).
    DOCX/PPTX/XLSX: MarkItDown (good for structured formats).
    Fallback: MarkItDown for PDF if pymupdf4llm not installed (PRESERVED).
    """
    if fmt == "pdf" and extractor == "tesseract":
        # Tesseract OCR via PyMuPDF OCR bridge
        import fitz
        doc = fitz.open(str(file_path))
        pages = []
        for page in doc:
            tp = page.get_textpage_ocr(language="eng", dpi=300, full=False)
            pages.append(page.get_text("text", textpage=tp))
        doc.close()
        return "\n\n".join(pages)
    elif fmt == "pdf" and pymupdf4llm is not None:
        mode = "layout" if _HAS_LAYOUT else "standard"
        print(f"  (pymupdf4llm mode: {mode})")
        try:
            md_text = pymupdf4llm.to_markdown(
                str(file_path),
                show_progress=False,
                write_images=False,
            )
        except ValueError as exc:
            # pymupdf/table.py raises ValueError when a table's
            # bounding-box list is empty (e.g. "min() iterable
            # argument is empty" at table.py:1534).  Exit with
            # code 3 so the pipeline router knows this extractor
            # failed and should try the next one in the chain.
            print(
                f"ERROR: pymupdf4llm crashed with ValueError: {exc}",
                file=sys.stderr,
            )
            print(
                "  This is a known pymupdf4llm table-detection bug.",
                file=sys.stderr,
            )
            print(
                "  Pipeline router should fall back to mineru or tesseract.",
                file=sys.stderr,
            )
            sys.exit(3)
        except Exception as exc:
            # Catch any other pymupdf4llm runtime error so the
            # process exits cleanly (non-zero) and the router can
            # trigger its fallback chain instead of seeing a raw
            # Python traceback.
            print(
                f"ERROR: pymupdf4llm failed unexpectedly: "
                f"{type(exc).__name__}: {exc}",
                file=sys.stderr,
            )
            sys.exit(3)
        return md_text
    elif MarkItDown is not None:
        md = MarkItDown()
        result = md.convert(str(file_path))
        return result.text_content
    else:
        print("ERROR: No text extractor available.")
        print("  PDF: pip install pymupdf4llm")
        print("  Other: pip install 'markitdown[all]'")
        sys.exit(1)


def clean_text(raw_text: str) -> str:
    """Clean common MarkItDown artifacts."""
    text = re.sub(r'\n{3,}', '\n\n', raw_text)
    lines = [line.rstrip() for line in text.split('\n')]
    return '\n'.join(lines).strip()


# ═══════════════════════════════════════════════════════════════════════════
# CONTEXT SUMMARY GENERATION (NEW)
# ═══════════════════════════════════════════════════════════════════════════

def parse_section_structure(md_text: str) -> list[dict]:
    """
    Parse section headings, line ranges, word counts, and image references.
    Returns list of section dicts.
    """
    sections = []
    lines = md_text.split('\n')
    current_section = None

    for line_num, line in enumerate(lines, 1):
        # Detect headings
        heading_match = re.match(r'^(#{1,6})\s+(.+)', line)
        if heading_match:
            # Save previous section
            if current_section:
                current_section["line_end"] = line_num - 1
                sections.append(current_section)

            # Start new section
            level = len(heading_match.group(1))
            heading_text = heading_match.group(2).strip()
            current_section = {
                "heading": heading_text,
                "level": level,
                "line_start": line_num,
                "line_end": None,
                "word_count": 0,
                "has_images": False,
                "image_refs": [],
            }

        # Count words and detect images in current section
        if current_section:
            current_section["word_count"] += len(line.split())
            # Detect image references
            img_refs = re.findall(r'\[([^\]]+)\]\(([^)]+\.(?:png|jpg|jpeg|gif|bmp))\)', line)
            if img_refs:
                current_section["has_images"] = True
                for label, path in img_refs:
                    # Extract figure number if present
                    fig_match = re.search(r'fig[\s-]?(\d+)', path, re.IGNORECASE)
                    if fig_match:
                        current_section["image_refs"].append(f"fig{fig_match.group(1)}")

    # Save last section
    if current_section:
        current_section["line_end"] = len(lines)
        sections.append(current_section)

    return sections


def detect_document_domain(md_text: str) -> str:
    """
    Auto-detect document domain from keyword frequencies.
    Returns one of: health_economics, clinical_trial, systematic_review,
    hta_regulatory, epidemiology, methodology, general.

    High-specificity override keywords are checked first. If any override
    keyword is found (case-insensitive substring), that domain wins
    immediately — bypassing the frequency scoring that can be skewed by
    generic clinical terms.

    S21/RC1: health_economics overrides take priority over hta_regulatory
    when BOTH have override matches.  hta_regulatory only wins in Phase 1
    when health_economics has zero override matches (i.e. the document is
    about regulatory process/HTA methodology without economic evaluation).
    """
    text_lower = md_text.lower()

    # ── Phase 1: Override keywords (high-specificity terms) ──
    # Collect ALL override matches first, then resolve priority.
    override_hits = {}  # domain -> list of matched keywords
    for domain, override_kws in DOMAIN_OVERRIDE_KEYWORDS.items():
        matched = [kw for kw in override_kws if kw.lower() in text_lower]
        if matched:
            override_hits[domain] = matched

    if override_hits:
        # S21/RC1: health_economics always wins over hta_regulatory when
        # both have override hits (documents about cost-effectiveness that
        # also mention HTA bodies are health_economics, not regulatory).
        if "health_economics" in override_hits:
            return "health_economics"
        # Only one domain matched, or hta_regulatory without health_econ
        best_domain = max(override_hits, key=lambda d: len(override_hits[d]))
        return best_domain

    # ── Phase 2: Frequency scoring (original logic) ──
    domain_scores = {}
    for domain, keywords in DOMAIN_KEYWORDS.items():
        score = sum(1 for kw in keywords if kw.lower() in text_lower)
        domain_scores[domain] = score

    # Return domain with highest score, or "general" if all zero
    max_domain = max(domain_scores.items(), key=lambda x: x[1])
    if max_domain[1] == 0:
        return "general"
    return max_domain[0]


def extract_key_terms(md_text: str, domain: str) -> list[str]:
    """Extract the top 5 most relevant keywords for the detected domain."""
    text_lower = md_text.lower()
    domain_keywords = DOMAIN_KEYWORDS.get(domain, [])

    # Count occurrences
    term_counts = {}
    for kw in domain_keywords:
        count = text_lower.count(kw.lower())
        if count > 0:
            term_counts[kw] = count

    # Return top 5
    sorted_terms = sorted(term_counts.items(), key=lambda x: x[1], reverse=True)
    return [term for term, count in sorted_terms[:5]]


def _is_institutional_header(text: str) -> bool:
    """Check if an H1 text is a generic institutional header, not a title.

    Uses case-insensitive substring matching against INSTITUTIONAL_HEADERS.
    Also rejects very short H1s (<=5 chars) which are usually acronyms or
    page numbers, not titles.
    """
    if not text or len(text) <= 5:
        return True
    text_lower = text.lower().strip()
    for header in INSTITUTIONAL_HEADERS:
        if header in text_lower:
            return True
    return False


def extract_title_from_md(md_text: str) -> Optional[str]:
    """Extract the first non-institutional H1 from markdown text.

    Scans the first 80 lines for H1 headings, skipping any that match
    INSTITUTIONAL_HEADERS (cover-page headers, ToC, etc.). Returns the
    first valid H1 text, or None if no valid H1 is found.

    This is a standalone helper used by both the pymupdf path
    (extract_title_authors) and the MinerU YAML frontmatter writer.
    """
    lines = md_text.split('\n')
    for line in lines[:80]:
        if line.startswith('# ') and not line.startswith('## '):
            h1_text = re.sub(r'^#+\s*', '', line).strip()
            if not _is_institutional_header(h1_text):
                return h1_text
    return None


def extract_title_authors(md_text: str) -> tuple[Optional[str], list[str]]:
    """
    Extract title (first non-institutional H1) and authors if present.
    Returns (title, authors_list).

    Skips institutional headers like "Health Technology Assessment" or
    "Statens legemiddelverk" which appear as H1 on cover pages but are
    not document titles. See INSTITUTIONAL_HEADERS constant.
    """
    title = extract_title_from_md(md_text)
    authors = []

    lines = md_text.split('\n')
    for line in lines[:80]:  # Expanded from 50 to 80 lines
        # Look for author line patterns
        if re.search(r'\bauthor', line, re.IGNORECASE):
            # Simple heuristic: names after "Authors:"
            author_match = re.search(r'authors?:?\s*(.+)', line, re.IGNORECASE)
            if author_match:
                author_str = author_match.group(1)
                # Split on common delimiters
                authors = re.split(r'[,;]\s*|\sand\s', author_str)
                authors = [a.strip() for a in authors if a.strip()]
                break

    return title, authors


def extract_abstract(md_text: str) -> Optional[str]:
    """
    Extract first 500 chars of text after an "Abstract" heading.
    Returns None if no abstract found.
    """
    lines = md_text.split('\n')
    in_abstract = False
    abstract_lines = []

    for line in lines:
        if re.match(r'^#{1,3}\s*abstract', line, re.IGNORECASE):
            in_abstract = True
            continue

        if in_abstract:
            # Stop at next heading
            if re.match(r'^#{1,3}\s', line):
                break
            abstract_lines.append(line)

    abstract_text = ' '.join(abstract_lines).strip()
    if abstract_text:
        abstract_text = re.sub(r'\s+', ' ', abstract_text)
        return abstract_text[:500]
    return None


def generate_context_summary(md_text: str, image_count: int) -> dict:
    """
    Generate context summary JSON from extracted markdown.
    No LLM — pure extractive/heuristic processing.
    """
    title, authors = extract_title_authors(md_text)
    abstract = extract_abstract(md_text)
    sections = parse_section_structure(md_text)
    domain = detect_document_domain(md_text)
    key_terms = extract_key_terms(md_text, domain)

    # Calculate total word count
    total_words = sum(s["word_count"] for s in sections)

    return {
        "title": title or "Unknown Title",
        "authors": authors,
        "abstract": abstract,
        "sections": sections,
        "total_sections": len(sections),
        "total_words": total_words,
        "total_images": image_count,
        "document_domain": domain,
        "key_terms": key_terms,
    }


# ═══════════════════════════════════════════════════════════════════════════
# YAML HEADER, IMAGE INDEX, MANIFEST
# ═══════════════════════════════════════════════════════════════════════════

def build_header(file_path: Path, fmt: str, page_count: Optional[int]) -> str:
    """Build YAML header block."""
    tool_map = {
        "pdf": "pymupdf4llm + PyMuPDF" if pymupdf4llm else "MarkItDown + PyMuPDF",
        "docx": "MarkItDown + python-docx",
        "pptx": "MarkItDown + python-pptx",
        "xlsx": "MarkItDown",
    }
    doc_type_map = {
        "pdf": "research_paper",
        "docx": "document",
        "pptx": "presentation",
        "xlsx": "spreadsheet",
    }

    header = "---\n"
    header += f"source_file: {file_path.name}\n"
    header += f"source_path: {file_path}\n"
    header += f"source_format: {fmt}\n"
    if page_count is not None:
        label = "slides" if fmt == "pptx" else "sheets" if fmt == "xlsx" else "pages"
        header += f"{label}: {page_count}\n"
    header += f"conversion_date: {datetime.now().strftime('%Y-%m-%d')}\n"
    header += f"conversion_tool: {tool_map[fmt]}\n"
    header += "fidelity_standard: verbatim (QC required)\n"
    header += f"document_type: {doc_type_map[fmt]}\n"
    header += "image_notes: pending\n"
    header += "---\n\n"
    return header


def build_image_index(entries: list[dict], images_rel: str) -> str:
    """Build image index table."""
    if not entries:
        return "## Image Index\n\nNo images extracted.\n\n"

    page_label = "Page/Slide"
    table = "## Image Index\n\n"
    table += f"| Fig | Description | File | {page_label} | Size | Source |\n"
    table += f"|-----|-------------|------|{'-' * len(page_label)}|------|--------|\n"

    for e in entries:
        page = e["page"] if e["page"] is not None else "-"
        table += (
            f"| {e['figure_num']} "
            f"| {e['description']} "
            f"| [{e['filename']}]({images_rel}/{e['filename']}) "
            f"| {page} "
            f"| {e['width']}x{e['height']} "
            f"| {e['source_format']} |\n"
        )

    table += "\n"
    return table


def write_image_manifest(entries: list[dict], images_dir: Path,
                         md_path: Path):
    """
    Write a JSON manifest of all extracted images with enhanced metadata.
    Used by prepare-image-analysis.py and generate-image-notes.md.
    """
    manifest = {
        "md_file": str(md_path),
        "images_dir": str(images_dir),
        "image_count": len(entries),
        "generated": datetime.now().isoformat(),
        "images": entries,
    }
    manifest_path = images_dir / "image-manifest.json"
    with open(manifest_path, "w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=2, ensure_ascii=False)
    print(f"  Image manifest: {manifest_path}")


def write_context_summary(summary: dict, output_dir: Path):
    """Write context summary JSON alongside the MD file."""
    summary_path = output_dir / "context-summary.json"
    with open(summary_path, "w", encoding="utf-8") as f:
        json.dump(summary, f, indent=2, ensure_ascii=False)
    print(f"  Context summary: {summary_path}")


# ═══════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(
        description="Convert PDF/DOCX/PPTX/XLSX to high-fidelity Markdown",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Supported formats: .pdf .docx .pptx .xlsx

Examples:
  python3 convert-paper.py paper.pdf
  python3 convert-paper.py report.docx -o reference-papers/report.md
  python3 convert-paper.py slides.pptx -i images/slides/
  python3 convert-paper.py data.xlsx --no-images

After conversion, run the QC pipeline:
  1. python3 qc-structural.py <output.md>
  2. python3 prepare-image-analysis.py <output.md>
  3. Claude subagent: generate-image-notes.md
  4. python3 validate-image-notes.py <output.md>
  5. Claude subagent: qc-content-fidelity.md (PDF only)
  6. Claude subagent: qc-final-review.md
        """
    )

    parser.add_argument("input_file", type=Path, help="Input file (PDF/DOCX/PPTX/XLSX)")
    parser.add_argument("--output", "-o", type=Path, default=None,
                        help="Output markdown path (default: <input>.md)")
    parser.add_argument("--images-dir", "-i", type=Path, default=None,
                        help="Images directory (default: images/<short-name>/)")
    parser.add_argument("--no-images", action="store_true",
                        help="Skip image extraction")
    parser.add_argument("--short-name", "-s", type=str, default=None,
                        help="Short name for filenames (default: from input name)")
    parser.add_argument("--extractor",
                        choices=["pymupdf4llm", "tesseract", "markitdown"],
                        default="pymupdf4llm",
                        help="Force extractor (default: pymupdf4llm for PDF, markitdown for other)")

    args = parser.parse_args()

    if not args.input_file.exists():
        print(f"ERROR: File not found: {args.input_file}")
        sys.exit(1)

    fmt = detect_format(args.input_file)
    short_name = args.short_name or args.input_file.stem.lower().replace(" ", "-")
    output_path = args.output or args.input_file.with_suffix(".md")
    images_dir = args.images_dir or (output_path.parent / "images" / short_name)
    page_count = get_page_count(args.input_file, fmt)

    page_label = "slides" if fmt == "pptx" else "sheets" if fmt == "xlsx" else "pages"
    print(f"Input:   {args.input_file}")
    print(f"Format:  {fmt.upper()}")
    print(f"Output:  {output_path}")
    if page_count:
        print(f"{page_label.capitalize()}: {page_count}")
    print(f"Images:  {images_dir}")
    print()

    # ── Step 1: Text extraction ──
    step_total = 3 if not args.no_images else 2
    extractor_name = args.extractor if fmt == "pdf" else "MarkItDown"
    print(f"[1/{step_total}] Extracting text with {extractor_name}...")
    raw_text = extract_text(args.input_file, fmt,
                            extractor=args.extractor)
    cleaned = clean_text(raw_text)
    print(f"  {len(cleaned):,} chars, {cleaned.count(chr(10)):,} lines")

    # ── Step 2: Image extraction ──
    image_entries = []
    if not args.no_images:
        print(f"[2/{step_total}] Extracting images...")
        if fmt == "pdf":
            import fitz  # Import here for caption detection
            image_entries = extract_images_pdf(args.input_file, images_dir, short_name, cleaned)
        elif fmt == "docx":
            image_entries = extract_images_docx(args.input_file, images_dir, short_name)
        elif fmt == "pptx":
            image_entries = extract_images_pptx(args.input_file, images_dir, short_name)
        elif fmt == "xlsx":
            print("  XLSX: no image extraction (spreadsheet format)")
        print(f"  {len(image_entries)} image(s) extracted")
    else:
        print(f"[2/{step_total}] Skipping image extraction (--no-images)")

    # ── Step 3: Assemble ──
    step_num = step_total
    print(f"[{step_num}/{step_total}] Assembling markdown...")

    try:
        images_rel = os.path.relpath(images_dir, output_path.parent)
    except ValueError:
        images_rel = str(images_dir)

    content = build_header(args.input_file, fmt, page_count)
    content += build_image_index(image_entries, images_rel)
    content += "---\n\n"
    content += cleaned
    content += "\n"

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(content, encoding="utf-8")

    # Write image manifest for the IMAGE NOTE subagent
    if image_entries:
        write_image_manifest(image_entries, images_dir, output_path)

    # Generate and write context summary (NEW)
    context_summary = generate_context_summary(cleaned, len(image_entries))
    write_context_summary(context_summary, output_path.parent)

    line_count = content.count('\n')
    print(f"\nDone. {line_count:,} lines written to {output_path}")

    # ── QC instructions ──
    print()
    print("=" * 60)
    print("NEXT STEPS (mandatory):")
    print("=" * 60)
    print(f"  1. python3 qc-structural.py {output_path}")
    print(f"  2. python3 prepare-image-analysis.py {output_path}")
    print(f"  3. Claude subagent: generate-image-notes.md")
    print(f"  4. python3 validate-image-notes.py {output_path}")
    if fmt == "pdf":
        print(f"  5. python3 extract-numbers.py {args.input_file} {output_path}")
        print(f"  6. Claude subagent: qc-content-fidelity.md")
        print(f"  7. Claude subagent: qc-final-review.md")
    else:
        print(f"  5. Claude subagent: qc-final-review.md")
        print(f"  (Content fidelity check skipped for {fmt.upper()} -")
        print(f"   MarkItDown is more reliable on structured formats)")
    print("=" * 60)


if __name__ == "__main__":
    main()
