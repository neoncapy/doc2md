---
name: convert-documents
description: "Convert PDF, DOCX, PPTX, or TXT files to high-fidelity Markdown with complete visual descriptions. Trigger: user asks to convert a document, read a PDF/DOCX/PPTX/TXT, or run the conversion pipeline."
---

<objective>
AUTHORITY: Global rule. ALL projects, ALL chats. Non-negotiable.

ZERO MISSING TEXT TOLERANCE: Every word, table cell, heading, and footnote from
the source document must appear in the converted markdown. No text may be silently
dropped, truncated, or omitted at any pipeline stage. If text is missing,
qc-structural.py and qc-content-fidelity.md must catch it — fix and re-run.

Pipeline version: v3.2.4 (2026-02-22) — chart/SmartArt LibreOffice render fallback; auto-organization workflow (--target-dir); visual reports; issue tracking; registry v3.1; table-collapse detection (R17); post-conversion cleanup (R18); automatic image index generation (R19); project-level testable image index (R20); registry image metadata (R21); agent-descriptions prompt generation (m2); image-index-overrides.json manual classification override (m3); metadata quality improvements (S16); near-black pixel-percentage pass (S16); MinerU tables/ directory support (S16); MinerU section heading integer-level fix (S16)
Changelog v3.2.4 (S16): metadata quality — dynamic section_context ratio, DOMAIN_OVERRIDE_KEYWORDS high-specificity terms, INSTITUTIONAL_HEADERS blocklist + extract_title_from_md() fallback, MinerU YAML title quote-escaping; near-black pixel-percentage pass (>95% pixels below brightness 15); MinerU tables/ directory scanning with table_ prefix + mineru_source field; _find_mineru_section_heading() now checks text_level integer field (MinerU uses int 1-4) before # prefix fallback; Issue 3 DOC-3 HTML table dupes closed (NOT A BUG — correct pandoc behavior for merged cells/colspan); Issue 4 metadata quality FIXED (3 bugs); MinerU 27/72 image gap RESOLVED (72 was bad reference count, MinerU's 27 is correct)
Changelog v3.2.3 (S15): MinerU fallback integration (_normalize_mineru_output, _trigger_mineru_fallback, _generate_image_index_from_mineru_manifest, Step 1c pre-QC index); MinerU models_config.yml patch (v3→v5 OCR det model + layoutreader 713MB); pymupdf4llm crash fix (removed import pymupdf.layout from convert-paper.py); R17 table-collapse FP fix (text-table guard in qc-structural.py, 19/20 FP eliminated); Issue 6 bare paths fixed (3 locations in convert-office.py); near-black detection (PIL+numpy mean<10 AND std<5, pdftoppm 300 DPI re-render); E2E validated: PPTX-1 (39 slides, 30 imgs), PPTX-2 (21 slides, 69 imgs+chart), DOCX-1 (244 tables, 10 imgs), PDF-1 (47pp, 27 MinerU imgs), PDF-2 (211pp, 129 imgs, 6 near-black)
Changelog v3.2.0 (S10-S11): vector render copy step (Step 9c, C1/C2 fix); 2 new decorative heuristics H9 color-block + H10 low-density badge; manifest-decorative subtract in display_count (M3 fix, Step 6c); dead code cleanup (dark-cover heuristic removed as subsumed by H9); convert-office.py bumped to 3.2.0

This skill orchestrates the unified document-to-Markdown conversion pipeline. The pipeline
operates in 2 tiers: Tier 1 (Python, zero LLM tokens) handles text extraction, image
extraction, structural QC, persona activation, and number extraction. Tier 2 (Claude API)
handles multi-expert IMAGE NOTEs, content fidelity QC, and final review. Output is enhanced
markdown with expert annotations.

UNIFIED PIPELINE (v3.0): PDF, PPTX, and DOCX all go through the same vision analysis
path (Steps 1-3, image notes, QC). The distinction between "full support" and "limited
support" for office formats is eliminated. PPTX and DOCX receive the same image extraction
and multi-expert IMAGE NOTE treatment as PDF.

Single entry point for all formats:

ALL FORMATS: ~/.claude/scripts/run-pipeline.py

run-pipeline.py handles format detection, extractor routing, checkpointing, registry
updates, and manifest discovery for ALL formats. For PPTX and DOCX it delegates Tier 1
extraction to convert-office.py internally, then runs the same Steps 2-3 (structural QC,
image analysis preparation) that PDF uses.

convert-office.py is still callable directly but run-pipeline.py is the preferred
entry point for all production conversions.
</objective>

<persona_quick_reference>
EXPERT PERSONA QUICK-REFERENCE (image analysis — Step 4 / generate-image-notes.md):

Abbreviations: STAT=Statistician, ECON=Health Economist, CLIN=Clinical Trialist,
MODEL=Model Architect, VIZ=Visualization Critic, REG=Regulatory/HTA Analyst,
EPI=Epidemiologist, MKT=Market Access Analyst

| Persona | Abbrev | Always Active For | Conditional Triggers |
|---|---|---|---|
| Statistician | STAT | kaplan-meier, forest-plot, tornado-diagram, scatter, histogram, box-plot, heatmap, funnel-plot, log_cumulative_hazard_plot, network-diagram | Any image with HR/OR/RR/p/CI values |
| Health Economist | ECON | tornado-diagram, scatter (CE plane), decision-tree | cost/QALY/ICER/budget keywords in context; line-chart with cost/CEAC/projection context; kaplan-meier with extrapolation |
| Clinical Trialist | CLIN | flow-chart (CONSORT) | kaplan-meier with trial groups; forest-plot (subgroup); bar-chart (AE/response); box-plot (clinical outcomes) |
| Model Architect | MODEL | decision-tree, flow-chart (model), schematic (model) | flow-chart variants with economic model context |
| Visualization Critic | VIZ | ALL non-decorative images (universal) | N/A — always active |
| Regulatory/HTA Analyst | REG | flow-chart (PRISMA) | forest-plot for HTA; tornado-diagram with submission context; scatter (CE plane) with WTP line; CONSORT flow-charts |
| Epidemiologist | EPI | (none unconditional) | kaplan-meier (population/registry context); scatter (population data); bar-chart (burden/prevalence); line-chart (time trends); pie-chart (disease distribution) |
| Market Access Analyst | MKT | (never unconditional — requires commercial/industry document context) | submission/dossier/reimbursement/payer documents only; activates for survival extrapolation, subgroup emphasis, CE plane, budget impact, tornado if drug cost excluded |

Full activation matrix (all image types x all personas): ~/.claude/scripts/generate-image-notes.md

NOTE: type_guess-based persona routing is PDF-only for regular extracted images.
For chart/SmartArt rendered via the LibreOffice fallback (PPTX only), convert-office.py
sets type_guess="chart" or type_guess="diagram", which activate persona routing in
prepare-image-analysis.py. All other PPTX/DOCX images get type_guess=null, routing
to visualization_critic only. The Claude vision agent (Step 4) must determine image
types from visual inspection for null entries.
</persona_quick_reference>

<quick_start>
FASTEST PATH — PDF:

```bash
python3 ~/.claude/scripts/run-pipeline.py \
  ~/Documents/papers/example-paper.pdf \
  -o ~/Documents/papers/example-paper.md \
  -i ~/Documents/papers/images/ \
  -s example-paper
```

FASTEST PATH — PPTX (full image extraction + vision analysis, same as PDF):

```bash
python3 ~/.claude/scripts/run-pipeline.py \
  ~/Documents/slides/example-deck.pptx \
  --output /tmp/out/example-deck.md
```

FASTEST PATH — DOCX (full image extraction + vision analysis, same as PDF):

```bash
python3 ~/.claude/scripts/run-pipeline.py \
  ~/Documents/reports/example-report.docx \
  --output /tmp/out/example-report.md
```

DIRECT convert-office.py (when bypassing the orchestrator for quick extraction only):

```bash
python3 ~/.claude/scripts/convert-office.py \
  ~/Documents/slides/example-deck.pptx \
  --output-dir ~/Documents/slides/
```

FASTEST PATH — WITH AUTO-ORGANIZATION (v3.1, --target-dir):

```bash
# Convert and auto-organize output into a project folder:
python3 ~/.claude/scripts/run-pipeline.py \
  ~/Downloads/Lecture1.pptx \
  --target-dir /path/to/output/L1/

# Result:
#   L1/_originals/Lecture1.pptx   (source moved here)
#   L1/Lecture1.md                (output)
#   L1/Lecture1_images/           (extracted images)
#   L1/Lecture1_manifest.json
#   L1/PIPELINE-REPORT-*.md       (visual report — always written)
```

DRY RUN (preview what would happen without doing it):

```bash
python3 ~/.claude/scripts/run-pipeline.py \
  ~/Downloads/file.pptx \
  --target-dir ~/Documents/project/L1/ \
  --dry-run
```

ORGANIZE ONLY (conversion already ran, just organize existing output):

```bash
python3 ~/.claude/scripts/run-pipeline.py \
  ~/Downloads/file.pptx \
  --target-dir ~/Documents/project/L1/ \
  --organize-only
```

GENERATE AGENT DESCRIPTIONS (m2 — structured Step 4 prompt file):

```bash
# Standalone (after conversion already ran):
python3 ~/.claude/scripts/run-pipeline.py \
  ~/Documents/papers/example-paper.pdf --agent-descriptions

# Combined with conversion and --target-dir:
python3 ~/.claude/scripts/run-pipeline.py \
  ~/Downloads/Lecture1.pptx \
  --target-dir ~/Documents/project/L1/ \
  --agent-descriptions

# Result adds:
#   L1/Lecture1-agent-descriptions.md  (Step 4 prompt file for vision subagent)
```

For EPUB: user must first run `ebook-convert input.epub output.htmlz` in Calibre,
then `pandoc output.htmlz -o output.md`.

run-pipeline.py handles its own registry update for all formats. No additional steps
required for Tier 1. Follow reported Claude steps (Tier 2) in order after the script
completes.
</quick_start>

<decision_gate>
EVALUATE BEFORE READING ANY DOCUMENT:

MD-FIRST READING RULE (enforced by PreToolUse hook querying conversion registry):
NEVER read raw PDF/document content. Always read converted markdown.
The only exceptions are:
1. No converted markdown exists yet (trigger conversion first)
2. Specific verification against source during QC
3. Visual analysis requiring direct image extraction

IF a PDF path is encountered:
-> Check for co-located .md file or query conversion registry
-> If converted .md exists: read the .md, NEVER the PDF
-> If no .md exists: invoke pipeline, THEN read the .md

IF document is PPTX or DOCX:
-> Check registry by SHA-256 hash first (hook enforces MD-first for all formats)
-> If registry hit: read the registered .md (regardless of where it lives)
-> If no registry match: check for co-located .md
-> If no .md exists: invoke run-pipeline.py (orchestrates convert-office.py internally)

IF document is TXT:
-> Check for co-located .md file
-> If none: invoke convert-office.py (wraps with YAML frontmatter)

IF document is XLSX or XLSM:
-> v3.1 BEHAVIOR (R13): run-pipeline.py handles XLSX/XLSM gracefully (no longer an error).
-> With --target-dir: moves Excel file to [target-dir]/_originals/ unchanged, logs INFO
   entry to CONVERSION-ISSUES.md, exits 0. No conversion attempted.
-> Without --target-dir: prints skip message, exits 0. No error.
-> XLSX->MD is lossy (formulas become strings, named ranges lost) — conversion is never done.
-> Recommended: use openpyxl to query the Excel file directly.
-> If MD is truly needed: export from Excel to CSV/TXT first, then convert the TXT.

IF document is EPUB:
-> Use Calibre EPUB -> HTMLZ -> pandoc -> MD path (manual first step by user)

NEVER bypass the pipeline for production conversions. Quick-reference conversions may
use Marker v2 fast path (no QC, no IMAGE NOTEs) when available.
NOTE: Marker v2 fast path is NOT currently available — planned future extractor.
</decision_gate>

<routing_rules>
FORMAT-TO-EXTRACTOR DECISION TABLE:

| Format | Condition | Extractor | Entry Point |
|--------|-----------|-----------|-------------|
| EPUB | always | Calibre ebook-convert -> pandoc | manual + pandoc |
| PPTX | production | python-pptx (recursive GROUP extraction) | run-pipeline.py (calls convert-office.py) |
| DOCX | production | pandoc (text) + python-docx (images) | run-pipeline.py (calls convert-office.py) |
| TXT | always | direct read + YAML frontmatter | run-pipeline.py (calls convert-office.py) |
| XLSX/XLSM | always | NOT CONVERTED — graceful skip (R13, v3.1) | run-pipeline.py exits 0; with --target-dir moves to _originals/ |
| PDF | digital (>= 50 chars/page) | pymupdf4llm | run-pipeline.py |
| PDF | scanned (< 50 chars/page) | Tesseract via PyMuPDF | run-pipeline.py --extractor tesseract |
| PDF | Tesseract fails | MinerU | convert-mineru.py |
| PDF | MinerU fails | Zerox VLM (last resort) | NOT AVAILABLE — planned future extractor |
| PDF | quick reference only | Marker v2 | NOT AVAILABLE — planned future extractor |

PPTX/DOCX PIPELINE (v3.0 UNIFIED):
run-pipeline.py now orchestrates the full pipeline for PPTX and DOCX, not just PDF.
After convert-office.py completes Tier 1 (text + image extraction), run-pipeline.py:
- Runs Step 2: qc-structural.py (structural QC gate)
- Runs Step 3: manifest discovery + prepare-image-analysis.py (image analysis prep)
- Writes registry entry (safety-net alongside convert-office.py's own write)
- Reports remaining Claude steps (image notes, final review) same as PDF

Manifest discovery (Step 3) uses a prioritized candidate list:
1. images_dir/image-manifest.json (PDF canonical)
2. output_dir/{input_stem}_manifest.json (office primary — what convert-office.py writes)
3. output_dir/{short_name}_manifest.json (office alternate)
First existing candidate wins. This replaces the old single hardcoded path that silently
skipped image analysis for all PPTX/DOCX.

Office manifest format: normalized in v3.0 to match convert-paper.py format. Fields:
figure_num, filename, width, height, file_path, section_context, page — all present
so prepare-image-analysis.py can consume office manifests without a crash.

SCAN DETECTION (PDF only):
Threshold: 50 characters per page average.
Method: Sample up to 10 evenly-spaced pages via fitz, compute average character count.
Below threshold = scanned document = OCR path.
At or above threshold = digital document = pymupdf4llm path.
If scan detection itself fails: WARN, assume digital, proceed with pymupdf4llm.
</routing_rules>

<convert_office_reference>
SCRIPT: ~/.claude/scripts/convert-office.py
Version: 3.2.4 | Pipeline version: 3.2.4

USAGE:

```bash
python3 ~/.claude/scripts/convert-office.py \
  <input-file> \
  [--output-dir <dir>] \
  [--skip-vision]
```

ARGUMENTS:
  input-file     Path to .pptx, .docx, or .txt file. Required.
  --output-dir   Output directory. Defaults to same directory as input file.
  --skip-vision  Accepted by the CLI parser but currently has no effect.
                 The flag is passed to converter functions but is not yet
                 implemented in any converter. IMAGE NOTE markers are always
                 written regardless of this flag. Reserved for future use.

OUTPUTS (per conversion):
  {basename}.md            - Draft markdown with YAML frontmatter (all formats)
  {basename}_images/       - Directory of extracted image files (PPTX/DOCX only)
  {basename}_manifest.json - Image manifest with metadata per image (PPTX/DOCX only)
  TXT output: {basename}.md only (no images dir, no manifest)

YAML FRONTMATTER fields by format (v3.0 — all required fields added):
  Common (all formats): title, source_file, source_format, conversion_tool,
    conversion_date, total_images, pipeline_version, document_type, fidelity_standard
  PPTX/DOCX additionally: images_directory
  PPTX additionally: total_slides
  TXT: common fields only (no images_directory, no total_slides)
  document_type values: "presentation" (PPTX), "document" (DOCX), "text" (TXT)
  fidelity_standard values: "visual_content" (PPTX), "text_content" (DOCX, TXT)
  Note: document_type and fidelity_standard are required by qc-structural.py
    check_header_block(). Missing these causes FAIL (exit 1). Added in v3.0.

MANIFEST FORMAT (v3.0 normalized — matches convert-paper.py format):
  Each image entry in {basename}_manifest.json now includes:
    figure_num, filename, width, height, file_path (absolute), section_context,
    nearby_text, page, type_guess (None), detected_caption (None)
  Top-level manifest also includes: md_file, images_dir, image_count, generated
  This normalization enables prepare-image-analysis.py to consume office manifests
  without a crash. Original fields (id, file, dimensions, slide) are preserved
  alongside the new normalized fields for backward compatibility.

REGISTRY:
  Writes to ~/.claude/pipeline/conversion_registry.json on
  success (same schema as run-pipeline.py entries). The PreToolUse hook will
  find the converted .md via the registry on subsequent read attempts.
  run-pipeline.py also writes a safety-net registry entry for office formats
  (dedup logic in update_registry() handles double-writes harmlessly).

PPTX CONVERSION DETAIL:
  Library: python-pptx
  Text extraction: per-slide, all text frames iterated, vertical tabs
    normalized, excessive blank lines collapsed.
  Image extraction: recursive GROUP shape traversal (captures nested images
    that flat iteration misses). Per-image outputs:
    - PICTURE shapes: extracted as raw blob, WMF/EMF converted to PNG
    - CHART/SmartArt shapes: rendered via LibreOffice fallback (see below);
      python-pptx cannot extract these as image blobs directly
  Image metadata per manifest entry:
    id, file, slide, content_type, size_bytes, dimensions [w,h],
    position {left/top/width/height as % of slide}, decorative (bool,
    true if < 5KB), source_shape, context {slide_title, nearby_text}
  Chart/SmartArt rendered PNG manifest entries (after LibreOffice fallback) include:
    source_format = "pptx_soffice_render"
    type_guess = "chart" (for chart shapes) or "diagram" (for SmartArt)
    Note: source_format only appears on rendered PNG entries. The initial
    file=None placeholder entries (before the fallback runs) do NOT have
    source_format — they only have id, file, type, slide, context fields.
  Markdown structure: ## Slide N: Title, slide text, then image refs as:
    <!-- IMAGE: fname | Size: WxH [| Decorative: yes] | Context: slide_title -->
    ![Image id](fname)
    <!-- CHART: id | Note: Chart rendered by PowerPoint, not extractable -->
    The "| Decorative: yes" field is present only when decorative=true (image < 5KB).
    Rendered chart/SmartArt PNGs use the same IMAGE comment format as regular images.

DOCX CONVERSION DETAIL:
  Text: pandoc (--wrap=none, produces clean markdown). Fallback to
    python-docx paragraph extraction if pandoc fails.
  Images: python-docx relationship traversal (doc.part.rels). Extracts
    all image relationships by content_type. WMF/EMF converted to PNG.
  Image placement: appended as ## Extracted Images section at end of .md.
  Image comment format: <!-- IMAGE: fname | Size: WxH [(decorative)] -->
    The "(decorative)" suffix is present only when decorative=true (image < 5KB).
    Note: DOCX format differs from PPTX — no "| Decorative: yes" pipe field
    and no "| Context: ..." field. Decorative is a parenthetical suffix instead.
  Image metadata per manifest entry: same schema as PPTX but slide=null,
    position=null (DOCX has no positional concept).
  Note: DOCX images are extracted by relationship ID, not paragraph position.
    Image order in manifest follows relationship traversal order.

TXT CONVERSION DETAIL:
  Encoding: tries UTF-8 first, falls back to latin-1.
  Output: YAML frontmatter + h1 title (derived from filename) + raw content.
  No images. No manifest. total_images: 0.

WMF/EMF CONVERSION (PPTX and DOCX):
  Step 1: PIL Image.open() -> save as PNG
  Step 2 (fallback): soffice --headless --convert-to png
  soffice path: /opt/homebrew/bin/soffice
  If both fail: warning logged, WMF/EMF kept as-is in images dir.

DECORATIVE THRESHOLD: images with size_bytes < 5000 are flagged decorative=true.
  These are typically logos, separators, or icon-scale assets.
  IMAGE NOTE agents should treat decorative=true images as low priority.

DEPENDENCY REQUIREMENTS:
  PPTX: python-pptx (pip install python-pptx)
  DOCX: python-docx (pip install python-docx) + pandoc (brew install pandoc)
  TXT: stdlib only
  Image conversion: Pillow (pip install Pillow) + soffice (optional fallback)
</convert_office_reference>

<invocation_protocol>
ENTRY POINT — PDF (all production PDF conversions):

```bash
python3 ~/.claude/scripts/run-pipeline.py \
  INPUT_FILE \
  -o OUTPUT_MD \
  -i IMAGES_DIR \
  -s SHORT_NAME \
  [--skip-numbers]  # for non-PDF formats
```

<!-- Full paths: see cross_project_rules table below -->
run-pipeline.py executes these Python steps sequentially:

Step 0: Format detection + scan detection + extractor routing
Step 1: convert-paper.py INPUT --extractor SELECTED -o OUTPUT -i IMAGES -s NAME
Step 2: qc-structural.py OUTPUT_MD
        -> FAIL = STOP. Fix and re-run.
        -> WARN = Python continues execution (not a hard stop), but Claude MUST investigate
           and fix all warned issues before proceeding. Re-run until PASS (exit 0).
        -> PASS = continue.
Step 3: prepare-image-analysis.py OUTPUT_MD
Step 6a: extract-numbers.py INPUT_PDF OUTPUT_MD (PDF only)

On completion, run-pipeline.py reports which Claude steps remain:

Step 4: generate-image-notes.md via Task subagent (multi-persona IMAGE NOTEs)
Step 5: validate-image-notes.py OUTPUT_MD --analysis-manifest MANIFEST --generate-spot-checks
Step 6b: qc-content-fidelity.md via Task subagent (PDF only, reads number-diff-report.json)
Step 7: qc-final-review.md via Task subagent (DIFFERENT agent from Step 6b, reads spot-checks.json)

ENTRY POINT — PPTX / DOCX / TXT (production, v3.0 UNIFIED):

```bash
python3 ~/.claude/scripts/run-pipeline.py \
  INPUT_FILE \
  [--output OUTPUT_MD]
```

run-pipeline.py detects the format and orchestrates the full pipeline:
Step 1: convert-office.py (text + image extraction, registry write)
        BUG-4 FIX: after Step 1, output_md is reconciled to match what
        convert-office.py actually wrote (uses input stem, not --output basename)
Step 2: qc-structural.py OUTPUT_MD (same gate as PDF pipeline)
        -> FAIL = STOP. Fix and re-run.
        -> WARN = Fix warned issues, re-run until PASS (exit 0).
        -> PASS = continue.
Step 3: prepare-image-analysis.py OUTPUT_MD
        Manifest discovery uses prioritized candidate list (3 paths checked).
        Auto-detect also finds {md_stem}_manifest.json alongside .md.
        --manifest arg can be passed explicitly if auto-detect fails.
Step 3b: (same as PDF) vision analysis when prepare-image-analysis.py produces output

After run-pipeline.py completes, the Claude steps that remain are:

Step A: generate-image-notes.md via Task subagent
        (if manifest has non-decorative images with file entries)
Step B: validate-image-notes.py (same as Step 5 in PDF pipeline)
Step C: qc-final-review.md via Task subagent (independent agent)

NOTE: qc-content-fidelity.md (number cross-check) is SKIPPED for office formats.
Office format text extraction (pandoc, python-pptx) is considered reliable enough
to skip number diffing. Apply content fidelity QC manually only if output quality
is suspect.

DIRECT convert-office.py CALL (advanced/debug — bypasses Steps 2-3):

```bash
python3 ~/.claude/scripts/convert-office.py \
  INPUT_FILE \
  [--output-dir OUTPUT_DIR] \
  [--skip-vision]
```

Use ONLY for quick extraction without QC, or debugging Tier 1 in isolation.
Production conversions MUST go through run-pipeline.py.

DIRECT convert-paper.py CALL (advanced/debug only):

```bash
python3 ~/.claude/scripts/convert-paper.py INPUT_FILE \
  --extractor [pymupdf4llm|tesseract|markitdown] \
  -o OUTPUT_MD -i IMAGES_DIR -s SHORT_NAME
```

NEVER invoke convert-paper.py directly for production conversions. run-pipeline.py
handles routing, cross-validation, checkpointing, and registry updates.

convert-paper.py METADATA IMPROVEMENTS (v3.2.4 — S16):

  section_context: find_section_for_page() now uses a dynamic ratio
    (total_lines / total_pages) instead of the hardcoded 45 lines/page
    constant. Eliminates incorrect section attribution on dense or
    sparse documents.

  document_domain: DOMAIN_OVERRIDE_KEYWORDS dictionary checked BEFORE
    the generic keyword scan. High-specificity terms (e.g. "NoMA",
    "ICER", "cost per QALY") trigger exact domain assignment without
    being overridden by generic matches lower in the priority chain.

  document_title: INSTITUTIONAL_HEADERS blocklist prevents common
    institutional headers (e.g. "Norwegian Medicines Agency",
    "NICE Technology Appraisal") from being captured as the document
    title. When the primary extractor returns a blocklisted string,
    extract_title_from_md() is called as a fallback to find the first
    meaningful heading in the converted markdown.

  MinerU YAML title: quote-escaping applied to the title field written
    into YAML frontmatter for MinerU-sourced documents. Prevents YAML
    parse errors when the title contains colons or special characters.

ANTI-FILTER RULE: ALL document text MUST go through Bash+Python. Claude NEVER outputs
bulk document text via Write/Edit tools. The Anthropic copyright filter blocks dense
text in API output. Bash+Python bypasses it.
</invocation_protocol>

<office_workflow>
COMPLETE OFFICE FORMAT WORKFLOW (PPTX / DOCX / TXT):

Stage 1 — CONVERSION (Tier 1, Python, zero tokens):

```bash
python3 ~/.claude/scripts/convert-office.py \
  /path/to/input.pptx \
  --output-dir /path/to/output/
```

Produces:
  input.md                  (draft markdown + YAML frontmatter)
  input_images/             (extracted image files)
  input_manifest.json       (image metadata)
  Registry entry written to ~/.claude/pipeline/conversion_registry.json

Stage 2 — STRUCTURAL QC (Python, free):
  When using run-pipeline.py: qc-structural.py is MANDATORY; run-pipeline.py always
  runs Step 2 unconditionally for all formats. Recommended for direct conversion
  calls (convert-office.py bypass mode) on PPTX/DOCX documents longer than 20
  slides / 10 pages.
  It catches encoding errors, malformed YAML frontmatter, and broken markdown
  syntax regardless of format.

```bash
python3 ~/.claude/scripts/qc-structural.py /path/to/input.md
```

  Exit 0 = PASS. Exit 1 = FAIL (fix before proceeding). Exit 2 = WARN (fix and re-run).

Stage 3 — IMAGE NOTES (Tier 2, Claude, Task subagent):
  Only run if manifest has images with file entries AND decorative=false.
  Use generate-image-notes.md prompt. Pass manifest path and images dir.
  Agent writes IMAGE NOTE blocks back into the .md file.

Stage 4 — IMAGE NOTE VALIDATION (Python):
```bash
python3 ~/.claude/scripts/validate-image-notes.py \
  /path/to/input.md \
  --analysis-manifest /path/to/images_dir/analysis-manifest.json \
  --generate-spot-checks
```
  Note: --analysis-manifest takes the ANALYSIS manifest (analysis-manifest.json,
  written by prepare-image-analysis.py to images_dir/), NOT the image manifest
  ({basename}_manifest.json written by convert-office.py or convert-paper.py).
  The flag name --analysis-manifest is correct as implemented in the script.

Stage 5 — FINAL REVIEW (Tier 2, Claude, independent Task subagent):
  Use qc-final-review.md prompt. MUST be a different agent from Stage 3.
  Reads spot-checks.json. Checks completeness, header, markdown rendering.

CHART / SMARTART RENDERING FALLBACK (PPTX only):
  python-pptx cannot extract native chart objects or SmartArt as image blobs.
  convert-office.py v3.1 adds a LibreOffice fallback to render these as PNGs.

  DETECTION:
    Charts:   shape.has_chart == True
    SmartArt: 3-tier detection —
              1. MSO_SHAPE_TYPE.IGX_GRAPHIC enum comparison (shape.shape_type == MSO_SHAPE_TYPE.IGX_GRAPHIC)
              2. MSO_SHAPE_TYPE.DIAGRAM enum comparison (shape.shape_type == MSO_SHAPE_TYPE.DIAGRAM)
              3. XML namespace check: "openxmlformats.org/drawingml/2006/diagram" in shape XML
              Tiers 1 and 2 use MSO_SHAPE_TYPE enum comparisons (not XML element search).
              Tier 3 is the XML-based fallback for edge cases not caught by enum checks.

  RENDERING (2-step via LibreOffice):
    Step 1: soffice --headless --convert-to pdf (PPTX to PDF)
            Note: soffice --convert-to png renders only slide 1 permanently.
            PDF intermediate is required to reach individual slides.
    Step 2: pdftoppm -png -r 300 -f N -l N (per-slide extraction from PDF)
            RENDER_DPI = "300" (string, not int — passed directly to pdftoppm CLI -r arg)
            Per-slide unique prefix: f"slide-{slide_idx:03d}" prevents filename collision.
    soffice path: shutil.which("soffice") first; falls back to /opt/homebrew/bin/soffice
    pdftoppm path: shutil.which("pdftoppm") first; falls back to /opt/homebrew/bin/pdftoppm (poppler)

  MANIFEST ENTRIES for rendered images:
    source_format = "pptx_soffice_render"
    type_guess    = "chart" (chart shapes) | "diagram" (SmartArt shapes)
    These values activate STAT, ECON, MODEL, VIZ, and other personas
    in prepare-image-analysis.py ACTIVATION_MATRIX.

  DEDUP AND BLANK DETECTION:
    SHA-256 deduplication and blank-image detection apply to rendered PNGs
    identically to regular extracted images.

  MULTI-CHART SLIDES:
    First chart/SmartArt on a slide: placeholder replaced with image ref.
    Subsequent charts on the same slide: cross-reference comment written
    pointing to the first rendered image. No duplicate image refs per slide.

  BLANK RENDERS:
    If a rendered PNG is detected as blank: the placeholder IS replaced with a
    failure comment (e.g. <!-- CHART: id — could not render (blank output) -->).
    The original placeholder text is NOT preserved as-is.
    Log warning only; conversion continues normally.

  ERROR HANDLING:
    All rendering errors are warnings only. Conversion never crashes due to
    chart rendering failure. If soffice or pdftoppm is unavailable, the
    original CHART comment placeholder is preserved.

  DEPENDENCIES:
    LibreOffice (soffice): /opt/homebrew/bin/soffice (brew install --cask libreoffice)
    pdftoppm (poppler):    /opt/homebrew/bin/pdftoppm (brew install poppler)
    Both must be installed for the fallback to activate.

SKIP RULES for office formats:
  SKIP: extract-numbers.py (PDF only)
  SKIP: qc-content-fidelity.md (PDF only)
  RECOMMENDED (direct convert-office.py only): qc-structural.py — MANDATORY when run through run-pipeline.py
  MANDATORY: generate-image-notes.md (if non-decorative images exist)
  MANDATORY: validate-image-notes.py (if IMAGE NOTEs were written)
  MANDATORY: qc-final-review.md (always)
</office_workflow>

<v31_organization>
AUTO-ORGANIZATION WORKFLOW (v3.1 — activated by --target-dir flag):

When --target-dir is NOT provided: pipeline behaves identically to v3.0 (R16 backward
compatibility). No _originals/, no CONVERSION-ISSUES.md, no PIPELINE-REPORT-*.md written.
Visual report IS printed to stdout but not saved to disk.

When --target-dir IS provided: pipeline runs Steps 7-13 after all conversion/QC steps
complete. Source file is never moved if conversion fails (R15 "move only after success").

NEW FLAGS (v3.1):
  --target-dir PATH, -t PATH   Target directory for organized output. Activates R1-R7.
  --force                      Bypass duplicate detection; re-convert even if hash found.
  --organize-only              Skip conversion (Steps 0-6); organize existing output only.
  --dry-run                    Print what WOULD happen; perform no file operations.

PIPELINE EXECUTION ORDER (v3.1 with --target-dir):

```
Step 0:  Format detection + scan detection + extractor routing
Step 1:  Conversion (text + images)
Step 1b: Cross-validation (PDF only)
Step 2:  Structural QC — GATE (FAIL = STOP; WARN = fix + re-run)
Step 3:  Prepare image analysis
Step 6a: Number extraction (PDF only)
--- v3.0 steps above; v3.1 organization steps below ---
Step 6c: R19 Image index generation (AUTOMATED — all formats; runs before organization)
         → writes [input_stem]-image-index.md alongside .md
         → filters decorative images; classifies substantive ones
         → M3 fix: when manifest data is available (PPTX/DOCX), images flagged as
           decorative in the manifest (size < 5KB during extraction, is_decorative_manifest=True)
           are subtracted from the per-page substantive display_count. This prevents
           double-counting when the manifest already marks icons/badges as decorative.
           Formula: display_count = total_detected - manifest_decorative_count (floor at 0).
           The image index header shows "Manifest-decorative images: N" when N > 0.
         → if --target-dir: manifest moves with .md in Step 9
Step 7:  R12 registry duplicate check (skip if same hash + same target-dir, unless --force)
Step 7:  R14 idempotent check (skip if source already in _originals/ with matching hash)
Step 7:  R5 pre-deletion verification (4 checks: .md exists, non-zero, has content, has YAML)
         Blocking failure (checks 1-3) → abort organization, log issue, exit code 2
         Advisory failure (check 4, YAML) → warning logged, proceed
Step 8:  R1 move source to [target-dir]/_originals/[filename] (atomic; cross-volume safe)
Step 9a: R2 place .md at [target-dir]/[input_stem].md
Step 9b: R3 place [input_stem]_images/, [input_stem]_manifest.json, and
         [input_stem]-image-index.md at [target-dir]/
Step 9c: C1/C2 FIX (PDF only) — Vector Render Copy
         After manifest update, vector renders stored in images/<slug>/ during
         chart/LibreOffice rendering (files matching patterns like
         *_vectorrender_*.png or *-vector-render.png) are copied to
         [target-dir]/[input_stem]_images/. Skips files already present.
         WHY: vector-heavy diagrams (forest plots, flowcharts, SmartArt) in
         medical PDFs are rendered by pymupdf/LibreOffice and written to
         images/<slug>/ (a different path than the main _images/ dir). Without
         this copy step, those renders are present in the source _images/ dir
         but missing from the organized [target-dir]/_images/. This step also
         merges vector render manifest entries into the target manifest, updating
         file_path references to the target directory.
Step 10: R4 cleanup intermediate files (.pipeline-checkpoint.json, .txt sidecars,
         MinerU working dirs, empty image dirs) — NEVER deletes /tmp/soffice-* (chart
         rendering cleans its own temp dirs)
Step 11: R10+R21 update registry with organized paths + image index metadata
Step 12: R7 generate visual report (includes image index summary) → printed to stdout
         + written to [target-dir]/PIPELINE-REPORT-[YYYYMMDD-HHMMSS].md
Step 13: R6 append any issues to [target-dir]/CONVERSION-ISSUES.md
```

OUTPUT DIRECTORY STRUCTURE after --target-dir run:

```
[target-dir]/
├── _originals/
│   └── source.pptx               (source file moved here — R1, R9)
├── source.md                      (converted output — R2)
├── source_images/                 (extracted images — R3)
│   └── image-001.png
├── source_manifest.json           (image manifest — R3)
├── source-image-index.md          (image index — R19; always created)
├── CONVERSION-ISSUES.md           (created on first issue — R6; not created if no issues)
└── PIPELINE-REPORT-20260221-110744.md  (always created — R7; includes image index summary)
```

R5 PRE-DELETION VERIFICATION (4 checks):
1. [target-dir]/[stem].md exists on disk
2. File size > 0 bytes
3. Contains at least one non-whitespace line
4. YAML frontmatter present (starts with --- on line 1) — advisory only
Checks 1-3 blocking: exit code 2 if any fail, source file untouched.
Check 4 advisory: warning logged, organization proceeds.

R6 CONVERSION ISSUE LOG ([target-dir]/CONVERSION-ISSUES.md):
Created on first issue only. Never overwritten (append only). Issue types:
  EXTRACTION_WARNING   — cross-validation flags, fallback extractor used
  QC_FAIL              — qc-structural.py exit code 1 (CRITICAL)
  VERIFICATION_FAIL    — R5 verification failed
  MOVE_FAIL            — R1 or R2 move failed (CRITICAL)
  DUPLICATE            — same file already organized (INFO)
  PARTIAL_CONVERSION   — conversion partial, some content may be missing
Severity: CRITICAL (stopped) | WARNING (continued with warning) | INFO (informational)

R7 VISUAL REPORT FORMAT (stdout + [target-dir]/PIPELINE-REPORT-*.md):
Unicode box-art frame (╔ ║ ╚ ═). Sections: CONVERSIONS, FILE MOVEMENTS, DELETED,
ISSUES, VERIFICATION, STATUS. Status: COMPLETE | COMPLETE WITH WARNINGS | FAILED.
Path truncation: last 60 chars with ... prefix. File sizes: bytes/KB/MB.
Status icons: ✓ (success), ⚠ (warning), ✗ (failure).

R11 VISUAL REPORT HELPER FUNCTIONS (run-pipeline.py, used by R7 report generator):
  human_readable_size(n): converts raw byte count to bytes/KB/MB string.
  truncate_path(p, max_len=60): truncates long paths to last max_len chars with … prefix.
  generate_visual_report(): calls both helpers to produce the R7 report output.
  These functions are internal implementation details of the R7 visual report.

R10 REGISTRY v3.1 FIELDS (added when --target-dir used):
  organized_source_path   → [target-dir]/_originals/[filename]
  organized_output_path   → [target-dir]/[stem].md
  organized_images_path   → [target-dir]/[stem]_images/
  target_dir              → the --target-dir value
  organized_at            → ISO 8601 timestamp (separate from converted_at)
  pipeline_version        → "3.1.0"
Migration rule: v3.0 entries without organized_* fields are NEVER modified.
When same hash found with no organized_* fields, a NEW entry is written alongside it.

R12 DUPLICATE DETECTION:
| Situation | Behavior |
|---|---|
| Same hash + same target-dir | Skip conversion; report ALREADY CONVERTED; update organized_at |
| Same hash + different target-dir | Proceed normally (intentional new location) |
| Different hash, same filename | Treat as new file; proceed normally |
| No hash match | Normal conversion |
--force flag bypasses all duplicate detection.

R13 EXCEL/XLSM BEHAVIOR (v3.1 — graceful skip, not error):
  With --target-dir: moves file to [target-dir]/_originals/; logs INFO; exit 0.
  Without --target-dir: prints skip message; exit 0.
  No conversion attempted. No registry entry written.

R14 IDEMPOTENCY: Running pipeline twice with same file + same --target-dir:
  Source already in _originals/ with matching SHA-256 → skip all organization; exit 0.
  .md already at target-dir → skip placement.
  Images already at target-dir → skip move.
  --force flag → re-organizes, overwrites .md, moves source again if needed.

R17 TABLE-COLLAPSE DETECTION (qc-structural.py v3.1):
qc-structural.py now detects collapsed multi-column tables in converted markdown.
Detection: counts numeric tokens per table row; flags tables where majority of data
rows have >= 1.5x numeric values vs declared column count. Single-column tables exempt.
Annotation: inserts HTML comment immediately after each flagged table:
  <!-- WARNING: Table may have collapsed columns. Original had N values
  but only M columns detected. Manual verification recommended. -->
Idempotent: re-running does not add duplicate comments.
Integration: flagged tables appear as WARN in qc-structural.py output and trigger a
WARNING entry in CONVERSION-ISSUES.md (when --target-dir is provided).
Test case: sample HTA exam PDF — 3-col and 4-col HTA tables
collapsed by pymupdf4llm; verified detected and annotated by R17.

R18 POST-CONVERSION CLEANUP AND PROJECT INTEGRATION (mandatory orchestration):
After EVERY conversion (with --target-dir), Claude MUST complete the full lifecycle
WITHOUT waiting for the user to ask. The Python pipeline (run-pipeline.py) handles Steps
1-3, 5-6. Step 4 (CLAUDE.md update) and Step 7 (report) are Claude orchestration steps.

MANDATORY 7-STEP LIFECYCLE:
  1. Convert (run-pipeline.py — existing)
  2. Move .md to correct project folder (R2 with --target-dir)
  3. Move original to [target-dir]/_originals/ (R1)
  4. Update project CLAUDE.md:
     → Add entry to Reference Files table for the new .md
     → Update Partial Components table status if applicable
     → Update folder structure comments to reflect new file
  5. Clean up source/working directory (delete ALL pipeline artifacts):
     → [stem]_images/ folder (if still at source after failed R3)
     → [stem]_manifest.json (if still at source)
     → .pipeline-checkpoint.json
     → Intermediate files (.docx created from .doc conversion, etc.)
     Scope: any pipeline artifact matching the specific input file's stem
  6. Verify cleanup: confirm source directory has zero pipeline artifacts
     from this conversion
  7. Report: single summary (converted, where moved, what cleaned up)

NEVER-ASK REQUIREMENT: Claude must complete Steps 4-7 automatically as part of
every conversion run. The user should NEVER need to request cleanup or CLAUDE.md
updates. Cleanup runs even if QC warnings were raised (warnings are non-blocking
for cleanup; only conversion failure blocks cleanup via R15).

SAFE CLEANUP PATTERNS:
  Safe to delete: [stem]_images/ at source, [stem]_manifest.json at source,
    .pipeline-checkpoint.json, .txt sidecars, MinerU working dirs
  NEVER delete: source file (moved, not deleted), .xlsx/.xlsm files,
    any file without a verified corresponding conversion output,
    /tmp/soffice-* (chart rendering cleans its own temp dirs),
    [stem]-image-index.md (permanent reference artifact — never delete)

R19 PER-FILE IMAGE INDEX GENERATION (AUTOMATED — Step 6c):
Runs automatically inside run-pipeline.py and convert-office.py after successful
conversion. No manual intervention required. Runs BEFORE file organization so the
manifest can be moved with the .md in Step 9.

Output file: [input_stem]-image-index.md (co-located with .md; moved to [target-dir]
when --target-dir is used).

What it does:
  - PDF: uses pymupdf (fitz) to detect images per page
  - PPTX: uses python-pptx to detect image shapes (including charts) per slide
  - DOCX: uses python-docx to detect inline and floating images per section
  - Extracts first meaningful line of text per page/slide as context (max 150 chars)
  - Classifies each page as SUBSTANTIVE or DECORATIVE using 10 heuristics:
    PDF path (run-pipeline.py):
      H1: images < 50x50 px → Decorative
      H2: page 1 with title/cover context → Decorative
      H3: last page with "thank"/"question" keyword → Decorative
      H4: same xref on >50% of pages → Decorative (watermark pattern)
      H5: images > 200x200 px with >20 chars text → Substantive
      H6: PPTX chart shapes → always Substantive (office path)
      H7: context contains figure keywords (Figure, Table, Diagram, Model) → Substantive
      H8: vector content via get_drawings() (threshold=50 with area guard >=1%;
          fallback at threshold=7 with area >=5%) → Substantive.
          BUG-3 fix: threshold raised from 7 to prevent styled table false positives.
      H9: color-block detection — images with < 32 unique colors (NEAREST resampling
          at 64x64) → Decorative. Catches solid-color backgrounds, simple gradients,
          corporate banners. Unique color counting uses NEAREST (not LANCZOS) to
          preserve exact pixel values; LANCZOS inflates counts via blending.
      H10: low-density badge detection — images with byte density < 0.15 B/px →
          Decorative. Catches simple logos, icons, badges with minimal visual info.
      Additional checks: near-uniform pixels (std < 5) or file size < 1KB → Decorative
          (FIX-2 blank detection). Journal branding (<100x100px, extreme aspect ratio
          >10:1, <5KB) → Decorative (FIX-3).
      Alpha handling: RGBA images no longer bypass H9/H10 heuristics. Only the
          std < 5 blank check is guarded by significant alpha (avoids false positives
          on transparent PNGs with uniform composited backgrounds).
    PPTX/DOCX path (convert-office.py):
      H1-H7 apply. Manifest-aware: decorative flag (< 5KB), duplicate detection,
      blank render detection. M1/M2/density heuristics (H9-H10) do NOT apply to
      office format blank detection (see DEDUP AND BLANK DETECTION note).
    → When uncertain: classify as SUBSTANTIVE (conservative — never miss a figure)
  - Documents with zero images produce a manifest with zero rows (not skipped)
  - Encrypted PDFs: manifest written with error note; conversion still succeeds
  - Failed index generation: conversion still succeeds; no image fields in registry

Manifest format: structured markdown with YAML-style header, page-by-page table
(all pages with images), substantive-only filtered table, and filtering summary.

TESTABLE IMAGE INDEX (R20 — project-level aggregate, --generate-testable-index):
Standalone operation that aggregates all per-file image indexes in a project into
a single reference file for study, review, or research sessions.

Usage:
  python3 ~/.claude/scripts/run-pipeline.py \
    --generate-testable-index /path/to/project-dir/

Output: [project-dir]/study-outputs/image-inventories/TESTABLE-IMAGE-INDEX.md
  - study-outputs/image-inventories/ is created automatically (mkdir -p) if missing
  - Running twice overwrites the previous index (complete regeneration, not append)

What it produces:
  - Scans all *-image-index.md files recursively in the project directory
  - Filters to SUBSTANTIVE images only
  - Groups entries by topic using keyword-based classification
  - Source category priority: CURRENT SLIDES > PREVIOUS SLIDES > LITERATURE > WORKING GROUP
  - Topic config: [project-dir]/.claude/config/image-index-topics.json
    (falls back to generic by-document grouping when config is absent)
  - "Why testable" descriptions: automated (template-based) by default;
    --agent-descriptions flag generates Sonnet-powered per-image prompt files (m2 — see below)

Topic config format (example for HTA project):
  HTA topics covered: Introduction to HTA, Costs & Costing, Quality of Life/HRQoL,
  Modelling, Cost-Effectiveness Analysis, Sensitivity & Uncertainty, Discounting,
  HTA & Policy Making, Transferability, Equity & Distribution, Theoretical Foundations.
  Full keyword map in IMAGE-INDEXING-REQUIREMENTS.md.

Proof of concept validated: 130 HTA PDFs → 5,238 inventory lines →
~95 high-value testable images across 11 topics. Reference files at:
  [your-project]/study-outputs/image-inventories/

R21 REGISTRY IMAGE INDEX FIELDS (7 new fields added per conversion):
When image indexing succeeds, these fields are added to the conversion registry entry
in ~/.claude/pipeline/conversion_registry.json:

  image_index_path          → absolute path to [stem]-image-index.md
  image_index_generated_at  → ISO 8601 timestamp of index generation
  total_pages               → total pages in source document
  pages_with_images         → pages containing at least 1 image
  total_images_detected     → raw count (including decorative)
  substantive_images        → count after decorative filtering
  has_testable_images       → boolean; true when substantive_images > 0

Migration rule: existing v3.0/v3.1 entries WITHOUT image fields are NEVER modified.
If image indexing fails: conversion succeeds; image fields are OMITTED from registry.

m2 — AGENT DESCRIPTIONS (--agent-descriptions):
Generates a structured prompt file for Claude vision subagents to use during Step 4
(image analysis / generate-image-notes.md). Replaces manual manifest inspection with
a ready-to-use prompt file that includes document context, image paths, classifications,
and persona assignments.

What it does:
  - Finds the analysis manifest and image manifest for the converted document
    (same discovery logic as Step 3 manifest auto-detect)
  - Reads classifications from the image index ([stem]-image-index.md),
    applying any overrides from image-index-overrides.json if present (m3)
  - Generates a structured prompt file with:
    • Document context (title, format, domain from YAML frontmatter)
    • Per-image entries: path, page number, SUBSTANTIVE/DECORATIVE classification,
      expert persona assignments (from prepare-image-analysis.py activation matrix)
    • Copy-paste ready instructions for the Step 4 vision subagent

Output: [stem]-agent-descriptions.md in the conversion output directory
  (moved to [target-dir]/ when --target-dir is used)

Usage (standalone — after conversion already ran):
  python3 ~/.claude/scripts/run-pipeline.py \
    file.pdf --agent-descriptions

Usage (combined with conversion):
  python3 ~/.claude/scripts/run-pipeline.py \
    file.pdf --target-dir /path/to/output --agent-descriptions

Usage (combined with --generate-testable-index):
  python3 ~/.claude/scripts/run-pipeline.py \
    --generate-testable-index /path/to/project/ --agent-descriptions

Availability: requires that conversion (Step 1) and image index generation (Step 6c)
have already run. If the image index is missing the flag exits with a clear error.
Model: Sonnet (classification + persona reasoning; does NOT require Opus vision).

m3 — IMAGE INDEX OVERRIDES (image-index-overrides.json):
Allows manual override of automatic SUBSTANTIVE/DECORATIVE classifications produced
by Step 6c (R19 image index generation). Useful when the heuristics mis-classify a
page (e.g. a small but critical diagram classified as DECORATIVE, or a decorative
header classified as SUBSTANTIVE).

File location: conversion output directory, alongside the image index.
  [target-dir]/image-index-overrides.json  (when --target-dir used)
  [same dir as .md]/image-index-overrides.json  (otherwise)

File format:
```json
{
  "overrides": [
    {"page": 5,  "classification": "SUBSTANTIVE", "reason": "Contains key model diagram"},
    {"page": 12, "classification": "DECORATIVE",  "reason": "Decorative chapter header"}
  ],
  "patterns": [
    {"pattern": "logo*",   "classification": "DECORATIVE",  "reason": "All logo images are decorative"},
    {"pattern": "fig_*",   "classification": "SUBSTANTIVE", "reason": "All figure images are substantive"}
  ]
}
```

Fields:
  overrides  → page-level overrides. "page" is the 1-based page / slide number.
               Applied first; exact match wins over patterns.
  patterns   → fnmatch-style glob patterns matched against image filenames.
               Applied after overrides. First matching pattern wins.
  reason     → human-readable provenance string (required; written to image index).

When overrides are applied:
  - During image index generation (Step 6c, run-pipeline.py): overrides are applied
    before the index is written; each overridden row includes a "MANUAL OVERRIDE"
    annotation and the reason string for full provenance.
  - During --agent-descriptions generation (m2): overrides are re-applied so the
    prompt file reflects the corrected classifications, not the raw heuristic output.

The image index records which pages were overridden. The registry is NOT updated
when overrides change (registry records heuristic counts; use the index for authoritative
classification after overrides).

To apply overrides to an already-generated index:
  1. Create or edit image-index-overrides.json in the output directory.
  2. Re-run: python3 ~/.claude/scripts/run-pipeline.py \
       file.pdf --generate-testable-index /path/to/project/
     (this regenerates the per-file index with overrides applied before aggregating)
  OR regenerate the index directly by re-running with --agent-descriptions (m2),
     which re-reads the manifest and re-applies overrides.
</v31_organization>

<extractor_selection>
The Extractor Selection Router lives in run-pipeline.py as the select_extractor() function.
This section covers PDF extractors only. Office format routing is in <routing_rules>.

DIGITAL CHAIN (avg_chars >= 50, covers 80%+ of documents):
1. pymupdf4llm via convert-paper.py (default)
2. After extraction: pdfplumber cross-validation (flags pages with >5% missing words)
3. If qc-structural.py reports table WARN/FAIL:
   a. Camelot re-extraction (called by QC, not router)
   b. Camelot fails -> pdfplumber table fallback
   c. Both fail -> MinerU fallback -> flag in checkpoint
4. Cross-validation flags written to checkpoint for QC consumption

SCANNED CHAIN (avg_chars < 50):
1. Tesseract OCR via PyMuPDF (page.get_textpage_ocr, fast, zero-dependency)
2. Tesseract unavailable or fails -> MinerU (~/envs/mineru, CPU-ONLY)
3. MinerU fails -> Zerox VLM (last resort, highest cost) [NOT AVAILABLE — planned future extractor]

CPU-ONLY RULE: MinerU MUST run CPU-only. No MPS/GPU. Previous GPU attempts caused kernel
panics (37.5GB RAM, WindowServer crash). This is non-negotiable.

AVAILABILITY CHECKS:
- Tesseract: shutil.which("tesseract") is not None
- MinerU: ~/envs/mineru directory exists
- Zerox: NOT AVAILABLE — planned future extractor (no convert-zerox.py on disk)

CROSS-VALIDATION (digital PDFs only):
After pymupdf4llm extraction, run-pipeline.py calls pdfplumber as secondary extractor.
Compare word sets per page. Flag pages where >5% of secondary words are missing from primary.
Flags are written to checkpoint JSON and consumed by qc-structural.py.

run-pipeline.py IMPROVEMENTS (v3.2.4 — S16):

  NEAR-BLACK DETECTION (additive — pixel-percentage pass):
    New 4th tier: if >95% of pixels have brightness below 15,
    the image is classified near-black and queued for pdftoppm
    300 DPI re-render. Additive to the existing 3-tier detection
    (file size, pixel std/mean/colors, PIL+numpy mean<10 AND std<5).
    RGBA handling added; PIL open errors caught and logged as warnings.

  MINERU TABLES/ DIRECTORY:
    _normalize_mineru_output() now scans BOTH images/ and tables/
    sub-directories in MinerU output. Table images get a table_ prefix
    (e.g. table_0.png, table_1.png). Each table image entry in the
    manifest includes a mineru_source field ("tables" vs "images") for
    downstream traceability.
    _generate_image_index_from_mineru_manifest() updated to handle
    table image entries correctly (path resolution uses tables/ subdir).

  MINERU SECTION HEADING FIX:
    _find_mineru_section_heading() now checks the text_level integer
    field FIRST. MinerU uses integer values 1-4 (not Markdown # prefix)
    to denote heading levels in its layout JSON. The previous
    implementation only checked for # prefix, missing all MinerU
    headings. Fallback to # prefix is preserved for robustness.
</extractor_selection>

<qc_interpretation>
QC PIPELINE: 3 passes, increasing depth. Each pass runs in a LOOP-UNTIL-ZERO pattern.

Note: QC subagent prompts use HIGH/MEDIUM/LOW severity labels (legacy convention from
extract-numbers.py). These map 1:1 to CRITICAL/MAJOR/MINOR. ALL levels must reach zero
regardless of naming.

╔═══════════════════════════════════════════════════════╗
║  QC LOOP-UNTIL-ZERO PROTOCOL (ALL QC STAGES)          ║
║                                                       ║
║  For each QC stage:                                   ║
║  1. Run QC agent/script                               ║
║  2. If issues found (ANY severity: CRITICAL, MAJOR,   ║
║     MINOR):                                           ║
║     a. Fix ALL issues (use Edit tool inline or        ║
║        Bash+Python for bulk fixes)                    ║
║     b. Re-run SAME QC agent/script                    ║
║     c. Repeat until ZERO issues at ALL severities     ║
║  3. Only proceed to next stage when current = ZERO    ║
║                                                       ║
║  ALL severities must reach zero. MINOR is NOT         ║
║  optional.                                            ║
║  No iteration cap. Loop runs until genuinely zero.    ║
║                                                       ║
║  Safety valve 1: SAME-ISSUE ESCAPE — if the SAME     ║
║  issue persists after 3 fix attempts, it is           ║
║  structural. Stop and escalate to user.               ║
║                                                       ║
║  Safety valve 2: TOKEN-BUDGET ESCAPE — if agent hits  ║
║  context limits, document ALL remaining issues for    ║
║  the next session. Do not silently drop issues.       ║
╚═══════════════════════════════════════════════════════╝

PASS 1 - STRUCTURAL QC (Python, free):
Script: ~/.claude/scripts/qc-structural.py
Exit codes: 0 = PASS, 1 = FAIL, 2 = WARN
Checks: YAML header completeness, section structure, table column consistency, reference
numbering, encoding errors, image index presence, manifest vs index consistency, image
file existence, markdown syntax.
Action on FAIL: fix reported issues, re-run until PASS.
Action on WARN: fix ALL warned issues, re-run until PASS (exit code 0). WARN is NOT a
pass-through. The loop continues until the script returns ZERO issues.
Action on PASS: proceed to next step.

PASS 2 - CONTENT FIDELITY (Claude + Python, PDF only):
Script (Python): ~/.claude/scripts/extract-numbers.py
Prompt (Claude): ~/.claude/scripts/qc-content-fidelity.md
Claude reads number-diff-report.json FIRST, then reviews ONLY flagged discrepancies.
ALL severity levels must be resolved to ZERO:
Severity: HIGH (missing critical data) -> verify PDF page + MD line, fix immediately.
Severity: MEDIUM (extra data) -> assess and resolve (remove if spurious, document if legitimate).
Severity: LOW (formatting) -> fix formatting differences, do not pass through.
After fixes: re-run extract-numbers.py + content fidelity review. Repeat until ZERO
discrepancies at ALL severities.
SKIP for DOCX, PPTX, XLSX, TXT (office format text extraction is reliable).

PASS 3 - FINAL REVIEW (Claude, independent agent):
Prompt: ~/.claude/scripts/qc-final-review.md
MUST be a different agent from Pass 2 (independent verification).
Reads spot-checks.json (pre-selected by validate-image-notes.py).
Checks: 5 pre-selected spot-check sections, image index completeness, header validation,
markdown rendering, persona analysis completeness, flagged image content review.
ALL issues found must be fixed and the review re-run until ZERO issues remain. No
severity is skippable.

IMAGE NOTE VALIDATION (Python, between Steps 4 and 7 / Stages 3 and 5):
Script: ~/.claude/scripts/validate-image-notes.py
Validates schema, persona blocks, cross-persona consistency, type-specific extensions.
Exit codes: 0 = PASS, 1 = FAIL (schema violations).
If FAIL: fix schema issues, re-run. Loop until exit code 0 (ZERO violations).
</qc_interpretation>

<failure_handling>
ERROR PATTERN: WARN vs FAIL

WARN = log to stderr, continue down the chain. Used for:
- Tesseract not installed (fall through to MinerU)
- MinerU venv missing (fall through to Zerox)
- Scan detection fails (assume digital, proceed)
- Registry write fails (conversion succeeded, registry is for future lookups)
- Registry is corrupt JSON (back up corrupt file, start fresh)
- pandoc subprocess raises an Exception (e.g. not installed, FileNotFoundError,
  timeout) — convert-office.py falls back to python-docx paragraph extraction.
  Note: pandoc returning a non-zero exit code does NOT trigger the fallback;
  it only logs a warning and uses stdout as-is (which may be partial or empty).
- soffice WMF/EMF conversion fails (keep original WMF/EMF in images dir)

FAIL = write to checkpoint, exit. Used for:
- Zerox fails (end of chain, no more fallbacks)
- Digital extraction produces 0 characters
- convert-paper.py returns non-zero exit code
- Text extraction fails completely (no recovery possible)
- Script not found at expected path (check ~/.claude/scripts/)
- convert-office.py: input file not found
- convert-office.py: unsupported format extension

PARTIAL FAILURE TABLE:

| Failure Type | Response | Recovery |
|---|---|---|
| Text extraction fails completely | FAIL document | No recovery, manual review |
| Single image description fails | Insert placeholder, continue | Retry that image later |
| QC step fails | Flag for manual review | Do not block other documents |
| Table extraction fails | Camelot -> pdfplumber -> MinerU -> flag | Chain through fallbacks |
| OCR fails on scanned page | MinerU -> Zerox VLM -> flag | Chain through fallbacks |
| Script not found | FAIL immediately | Verify scripts exist in ~/.claude/scripts/ |
| pandoc unavailable (DOCX) | Fall back to python-docx paragraph extraction | Quality may be lower |
| WMF/EMF conversion fails | Keep as WMF/EMF, log warning | Manual screenshot if image critical |
| PPTX chart shape (soffice/pdftoppm available) | LibreOffice fallback render | PNG added to manifest with source_format=pptx_soffice_render |
| PPTX chart shape (soffice/pdftoppm unavailable) | Preserve CHART comment placeholder, log warning | Install LibreOffice + poppler, re-run |
| PPTX chart render returns blank PNG | Replace placeholder with failure comment (e.g. <!-- CHART: id — could not render (blank output) -->), log warning | Manual screenshot from PowerPoint if content is critical |

PLACEHOLDER FORMAT for failed images:
<!-- IMAGE DESCRIPTION FAILED: figure_N.png -- manual review required -->

CHECKPOINT RECOVERY (PDF pipeline only):
Pipeline state is persisted in checkpoint JSON per document.
On crash at any step, checkpoint shows: completed steps, in-progress step, last completed index.
Resume from where checkpoint indicates. Do not re-run completed steps.
convert-office.py does not use checkpointing (single-call design). On failure, re-run the
full command. It is fast enough to restart without penalty.

FILTER RECOVERY (Anthropic copyright filter):
1. STOP immediately. Context is poisoned.
2. Do NOT retry in same agent/context.
3. Launch FRESH subagent (new Task tool call).
4. Fresh agent reads MD file, continues from where it stopped.
</failure_handling>

<anti_patterns>
- Reading raw PDF/DOCX/PPTX content instead of converted markdown (use MD-first rule)
- Invoking convert-paper.py directly for production conversions (use run-pipeline.py)
- Invoking convert-office.py directly for production conversions (use run-pipeline.py,
  which runs Steps 2-3 after convert-office.py; direct call skips QC and image analysis)
- Using Write/Edit tools for bulk document text (use Bash+Python, anti-filter rule)
- Using MPS/GPU for MinerU (CPU-only, kernel panic risk)
- Skipping structural QC before Claude QC passes (Step 2 must PASS before proceeding)
- Using the same agent for content fidelity QC and final review (must be independent)
- Re-running completed pipeline steps after checkpoint recovery (resume, do not restart)
- Copying pipeline scripts into project directories (centralized in ~/.claude/scripts/ only)
- Skipping IMAGE NOTE validation between image note generation and final review
  (schema violations propagate)
- Retrying in the same context after Anthropic copyright filter triggers (context is poisoned)
- Running qc-content-fidelity.md on office format output (PDF-only step, skip it)
- Running extract-numbers.py on office format output (PDF-only step, skip it)
- Ignoring rendered chart/SmartArt PNGs in PPTX output (they are real image files
  requiring IMAGE NOTEs; treat source_format="pptx_soffice_render" entries the same
  as regular extracted images in the manifest)
- Manually screenshotting charts when soffice+pdftoppm are available (use the
  LibreOffice render fallback instead; manual screenshots are only needed when
  the fallback is unavailable or produces a blank render)
- Treating decorative=true images as requiring IMAGE NOTEs (low priority, skip or
  write minimal notes)
- Using --target-dir without completing the R18 CLAUDE.md update step (Step 4 is
  mandatory — add entry to Reference Files table, update Partial Components if
  applicable, update folder structure; these are Claude orchestration steps, not Python)
- Asking the user "should I clean up?" after a conversion with --target-dir
  (R18 cleanup is non-negotiable and always runs automatically)
- Globbing /tmp/soffice-* for cleanup in run-pipeline.py (chart rendering cleans
  its own temp dirs inside convert-office.py; R4 cleanup must NOT touch these)
- Relying on the visual report stdout output as the only record (PIPELINE-REPORT-*.md
  is always written to disk with --target-dir for recovery after terminal closes)
- Treating R17 WARNING comments (collapsed table annotation) as errors to be removed
  (they are informational flags for manual review — leave them in the .md)
- Skipping --dry-run when uncertain about organization logic (dry-run costs nothing
  and shows exactly what will happen before any files are moved)
- Deleting [stem]-image-index.md during cleanup (it is a permanent reference artifact;
  NEVER delete it; it is listed alongside .md and images in the output structure)
- Treating image index generation as optional (R19 runs automatically for ALL formats;
  even text-only documents produce a zero-row manifest — this is correct behavior)
- Confusing Step 6c (image index, AUTOMATED) with Steps 4/5/6b (IMAGE NOTEs + QC,
  MANUAL/AGENT); Step 6c needs no user or agent action — it runs inside run-pipeline.py
- Running --agent-descriptions before conversion and image index generation have completed
  (the flag requires both [stem]_manifest.json and [stem]-image-index.md to exist;
  run conversion first, THEN add --agent-descriptions as a standalone call if needed)
- Editing the image index .md directly to change SUBSTANTIVE/DECORATIVE classifications
  (edits are overwritten on re-run; use image-index-overrides.json instead — m3)
- Omitting the "reason" field from image-index-overrides.json entries (it is required;
  it is written into the image index as provenance metadata)
- Expecting the conversion registry to update when overrides change (registry stores
  heuristic counts computed at conversion time; overrides affect the index only, not
  the registry fields total_images_detected / substantive_images)
</anti_patterns>

<agent_definitions>
<!-- NOTE: tag kept as agent_definitions for backward compatibility.
     Equivalent to subagent_configs in other skills. -->
For full agent definitions, YAML configs, output format, and parallelization rules, see:
~/.claude/skills/convert-documents/references/agent-definitions.md

For the 8 expert persona definitions, activation matrix, IMAGE NOTE schema, and
worked examples, see:
~/.claude/scripts/generate-image-notes.md
</agent_definitions>

<cross_project_rules>
SCRIPT LOCATIONS (centralized, NEVER project-level copies):
All pipeline scripts live in ~/.claude/scripts/ ONLY.

| Script | Full Path | Lines (v3.2.4) |
|--------|-----------|----------------|
| run-pipeline.py | ~/.claude/scripts/run-pipeline.py | 7,749 |
| convert-office.py | ~/.claude/scripts/convert-office.py | 3,122 |
| convert-paper.py | ~/.claude/scripts/convert-paper.py | 1,329 |
| qc-structural.py | ~/.claude/scripts/qc-structural.py | 1,211 |
| prepare-image-analysis.py | ~/.claude/scripts/prepare-image-analysis.py | 634 |
| convert-mineru.py | ~/.claude/scripts/convert-mineru.py | 228 |
| validate-image-notes.py | ~/.claude/scripts/validate-image-notes.py | — |
| extract-numbers.py | ~/.claude/scripts/extract-numbers.py | — |
| generate-image-notes.md | ~/.claude/scripts/generate-image-notes.md | — |
| qc-content-fidelity.md | ~/.claude/scripts/qc-content-fidelity.md | — |
| qc-final-review.md | ~/.claude/scripts/qc-final-review.md | — |

CONVERSION REGISTRY (shared across all projects):
Path: ~/.claude/pipeline/conversion_registry.json
Schema per entry (v3.0 base fields): source_hash (sha256), source_path, output_path,
extractor_used, converted_at (ISO 8601), pipeline_version.
v3.1 additional fields (present only when --target-dir was used):
  organized_source_path   → [target-dir]/_originals/[filename]
  organized_output_path   → [target-dir]/[stem].md
  organized_images_path   → [target-dir]/[stem]_images/
  target_dir              → the --target-dir path used
  organized_at            → ISO 8601 timestamp for the organization step
  pipeline_version        → "3.1.0"
Migration rule: v3.0 entries without organized_* fields are NEVER modified.
When same hash found with no organized_* fields, a NEW entry is appended alongside it.
Updated by run-pipeline.py after successful conversion (all formats: PDF, PPTX, DOCX, TXT).
Updated by convert-office.py after successful PPTX/DOCX/TXT conversion (primary write).
Updated by run-pipeline.py (safety-net) for office formats orchestrated via it (dedup handled).
Queried by the PreToolUse hook to enforce MD-first reading for ALL formats.
Hook registry lookup: v3.0 adds SHA-256 registry lookup for DOCX/PPTX/XLSX
(same jq query as PDF). Log event: "office-registry-hit". Allows reading the
.md regardless of where it lives (not just co-located with the source file).
This matters when source is in ~/Downloads but output .md is in a project folder.

PDF ARCHIVE (central):
Location: ~/Documents/pdf-archive/
Rule: symlink original PDFs from project directories to the central archive.

PER-PROJECT USAGE:
Each project that uses the pipeline creates an images/ directory for extracted visuals.
Output markdown goes into the project's own directory structure.
No pipeline scripts are ever copied into project directories.

YAML FRONTMATTER (mandatory on all converted files):
PDF conversions record: source_file, source_path, source_format, pages, conversion_date,
conversion_tool, fidelity_standard, document_type, document_domain, image_notes status,
persona_analysis status, persona_version, flagged_images.
Office format conversions (convert-office.py v3.0) record:
  Common (all formats): title, source_file, source_format, conversion_tool,
    conversion_date, total_images, pipeline_version, document_type, fidelity_standard.
  PPTX/DOCX additionally: images_directory.
  PPTX additionally: total_slides.
  TXT: common fields only (no images_directory, no total_slides).
  (document_type and fidelity_standard added in v3.0 — required by qc-structural.py)
Both satisfy NICE 2024 position statement requirements for transparency, reproducibility,
and AI disclosure.
</cross_project_rules>

<reference_index>
Reference files are at: ~/.claude/skills/convert-documents/references/
- ~/.claude/skills/convert-documents/references/agent-definitions.md
When passing reference file paths to agents, use the full absolute paths above.
</reference_index>

<success_criteria>
PDF CONVERSION COMPLETE when ALL of the following are true:

1. Structural QC PASS: qc-structural.py exits with code 0 (PASS). WARN (exit 2) must
   be looped until resolved to PASS.
2. IMAGE NOTEs written: all extracted images have multi-expert persona analysis blocks
   in the markdown.
3. IMAGE NOTE validation PASS: validate-image-notes.py exits with code 0 (schema valid).
4. Content fidelity QC PASS: qc-content-fidelity agent reports ZERO issues at ALL
   severities (HIGH, MEDIUM, LOW) (PDF only).
5. Final review PASS: qc-final-review agent (independent from content fidelity agent)
   confirms completeness.
6. Registry updated: conversion_registry.json contains entry with source_hash, paths,
   extractor, timestamp.
7. Checkpoint marked complete: checkpoint JSON shows all steps completed with no
   in-progress flags.

OFFICE FORMAT CONVERSION COMPLETE when ALL of the following are true:

1. convert-office.py exits cleanly: .md written (all formats), plus _images/ and
   _manifest.json (PPTX/DOCX only).
2. Registry updated: conversion_registry.json contains entry for the source file.
3. Structural QC PASS (if run): qc-structural.py exits with code 0.
4. IMAGE NOTEs written (if non-decorative images exist): all non-decorative images have
   IMAGE NOTE blocks in the markdown.
5. IMAGE NOTE validation PASS (if IMAGE NOTEs written): validate-image-notes.py exits
   with code 0.
6. Final review PASS: qc-final-review agent confirms completeness.
7. Charts/SmartArt covered (PPTX only): slides with native chart or SmartArt shapes
   have been rendered via LibreOffice fallback (source_format="pptx_soffice_render"
   entries in manifest) OR, if soffice/pdftoppm unavailable, CHART comment placeholders
   are present and noted in final review. All rendered chart PNGs have IMAGE NOTEs
   written by the vision agent. Blank renders are flagged in final review.

If any criterion is not met, the conversion is NOT complete. Do not mark as done.

WITH --target-dir (v3.1 organization complete) when ALL of the following are true:

1. All base conversion criteria above are met.
2. Source file is at [target-dir]/_originals/[filename] (not at original path).
3. Converted .md is at [target-dir]/[input_stem].md.
4. Images directory is at [target-dir]/[input_stem]_images/ (if images were extracted).
5. Manifest JSON is at [target-dir]/[input_stem]_manifest.json (if images were extracted).
6. Image index is at [target-dir]/[input_stem]-image-index.md (R19; always present;
   zero-row manifest for documents with no images is acceptable).
7. Zero pipeline artifacts remain at the source location (no [stem]_images/, no
   [stem]_manifest.json, no .pipeline-checkpoint.json from this conversion).
8. PIPELINE-REPORT-*.md written at [target-dir]/ with COMPLETE or COMPLETE WITH WARNINGS
   (includes image index summary section — pages scanned, substantive vs decorative count).
9. CONVERSION-ISSUES.md exists at [target-dir]/ only if issues occurred.
10. Registry updated with organized_* fields AND image index fields (R21):
    organized_source_path, organized_output_path, organized_at, pipeline_version: "3.1.0",
    image_index_path, image_index_generated_at, total_pages, pages_with_images,
    total_images_detected, substantive_images, has_testable_images.
11. Project CLAUDE.md updated: Reference Files table has entry for the new .md;
    Partial Components table updated if applicable; folder structure updated if applicable.
    (This is Claude's responsibility, not the Python pipeline's.)
12. Summary reported to user: what was converted, where it was moved, what was cleaned up,
    and image index stats (substantive images found).
</success_criteria>
