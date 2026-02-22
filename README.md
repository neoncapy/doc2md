## Summary
World-class README for the doc2md open-source document conversion pipeline. Covers architecture, installation, usage, Claude Code integration, and all component files.

---

<div align="center">

# doc2md

**High-fidelity document-to-Markdown conversion pipeline for Claude Code**

Convert PDF, DOCX, and PPTX files to structured Markdown with image extraction,
multi-stage quality control, and LLM-ready image analysis preparation.

[![Python 3.8+](https://img.shields.io/badge/python-3.8%2B-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Claude Code](https://img.shields.io/badge/Claude_Code-compatible-blueviolet.svg)](https://docs.anthropic.com/en/docs/claude-code)
[![Code style: black](https://img.shields.io/badge/code%20style-black-000000.svg)](https://github.com/psf/black)

[Quick Start](#quick-start) | [Architecture](#architecture) | [Usage](#usage) | [Claude Code Integration](#claude-code-integration) | [Configuration](#configuration)

</div>

---

## Why doc2md?

Anthropic's copyright filter blocks most direct PDF reads in Claude Code.
Even when reads succeed, raw PDF parsing loses tables, headings, and images.
This pipeline solves both problems:

- **Zero-token Python tier** extracts text and images with full structural fidelity
- **Optional LLM tier** generates expert image descriptions using 8 specialist personas
- **Multi-stage QC** catches table collapse, heading hierarchy errors, and missing content
- **SHA-256 registry** tracks every conversion, preventing duplicate work

The result: Markdown files that Claude Code can read, reason about, and reference
with full access to every word, table, heading, and figure from the source document.

---

## Features

| Feature | Description |
|---|---|
| **Unified router** | Single entry point handles PDF, DOCX, PPTX, and TXT |
| **Multi-extractor PDF** | pymupdf4llm (default), pdfplumber (cross-validation), MinerU (complex layouts) |
| **Office conversion** | DOCX via pandoc + python-docx; PPTX via python-pptx with recursive group shape extraction |
| **Chart rendering** | LibreOffice &rarr; PDF &rarr; pdftoppm at 300 DPI for SmartArt and embedded charts |
| **Image deduplication** | SHA-256 hashing skips duplicate images across pages |
| **Blank detection** | 3-tier detection: file size, pixel statistics, near-black analysis |
| **Per-image classification** | 8-heuristic chain classifies each image as substantive or decorative |
| **Vector content detection** | pymupdf `get_drawings()` identifies diagrams, SmartArt, shape-based figures |
| **Structural QC engine** | Automated checks for table collapse, heading hierarchy, YAML metadata, encoding errors |
| **Persona activation matrix** | Maps 24+ image types to 8 expert personas for targeted LLM analysis |
| **Conversion registry** | JSON registry with SHA-256 hashes, image metadata, conversion timestamps |
| **Image indexing** | Per-file and project-level testable image indexes |
| **Claude Code hook** | Enforces "never read raw PDF" policy at the tool level |
| **MinerU fallback** | Auto-switches to MinerU when cross-validation failure rate exceeds 40% |
| **DOCX table styling** | Professional styling for pandoc-generated Word documents |

---

## Architecture

```
                              ┌─────────────────────────────┐
                              │       run-pipeline.py        │
                              │      (unified router)        │
                              └──────────┬──────────────────┘
                                         │
                    ┌────────────────────┼────────────────────┐
                    ▼                    ▼                    ▼
              ┌──────────┐      ┌──────────────┐      ┌──────────┐
              │   PDF     │      │    DOCX      │      │   PPTX   │
              └────┬─────┘      └──────┬───────┘      └────┬─────┘
                   │                   │                    │
 ═══════════════════════════════════════════════════════════════════
  TIER 1: Python (zero LLM tokens)
 ═══════════════════════════════════════════════════════════════════
                   │                   │                    │
                   ▼                   ▼                    ▼
          ┌────────────────┐  ┌────────────────┐  ┌────────────────┐
          │ convert-paper  │  │ convert-office  │  │ convert-office  │
          │  pymupdf4llm   │  │  pandoc+docx    │  │  python-pptx   │
          └────────┬───────┘  └────────┬───────┘  └────────┬───────┘
                   │                   │                    │
                   └───────────┬───────┘────────────────────┘
                               │
                   ┌───────────▼───────────┐
                   │  Step 1b: Cross-Val   │  (PDF only: pdfplumber)
                   │  Step 2:  Structural  │  (QC gate — must PASS)
                   │  Step 3:  Image Prep  │  (persona activation)
                   │  Step 6c: Image Index │  (SUB/DEC classification)
                   └───────────┬───────────┘
                               │
 ═══════════════════════════════════════════════════════════════════
  TIER 2: Claude (LLM — optional, manual)
 ═══════════════════════════════════════════════════════════════════
                               │
                   ┌───────────▼───────────┐
                   │  Step 4: IMAGE NOTEs  │  (8 expert personas)
                   │  Step 5: Content QC   │  (fidelity check)
                   │  Step 6: Final Review │  (human-in-the-loop)
                   └───────────────────────┘
```

**Tier 1** runs entirely via Python. No API calls, no tokens consumed.
It extracts text, images, and metadata; runs structural QC; and prepares
the analysis manifest that tells Tier 2 which expert personas should
examine each image.

**Tier 2** is optional and uses Claude's vision capabilities to generate
multi-expert image descriptions. This tier is invoked manually through
Claude Code's agent system or the included skill definition.

---

## Quick Start

### Prerequisites

```bash
# Core dependencies
pip install pymupdf pymupdf4llm pdfplumber python-docx python-pptx Pillow numpy

# Pandoc (required for DOCX text extraction)
brew install pandoc        # macOS
sudo apt install pandoc    # Ubuntu/Debian

# Optional: LibreOffice (for chart/SmartArt rendering in PPTX)
brew install --cask libreoffice    # macOS

# Optional: MinerU (for complex/scanned PDFs)
# See https://github.com/opendatalab/MinerU for installation
```

### Install

```bash
git clone https://github.com/orangefineblue/doc2md.git
cd doc2md

# Copy scripts to your preferred location
cp scripts/*.py ~/.local/bin/    # or any directory in your PATH
```

### Run

```bash
# Convert a PDF
python3 scripts/run-pipeline.py paper.pdf -o paper.md -i images/

# Convert a DOCX
python3 scripts/run-pipeline.py report.docx -o report.md

# Convert a PPTX
python3 scripts/run-pipeline.py slides.pptx -o slides.md

# Convert with organized output directory
python3 scripts/run-pipeline.py paper.pdf --target-dir ./converted/
```

---

## Usage

### Basic Conversion

The unified router (`run-pipeline.py`) auto-detects file format and selects
the appropriate extractor:

```bash
python3 run-pipeline.py <input-file> [options]
```

| Option | Description |
|---|---|
| `-o`, `--output` | Output markdown file path |
| `-i`, `--images` | Image output directory |
| `-s`, `--short-name` | Short name for file references |
| `--target-dir` | Organized output directory (moves source to `_originals/`) |
| `--force-extractor` | Override extractor selection (pymupdf4llm, tesseract, mineru) |
| `--skip-xval` | Skip cross-validation step |
| `--dry-run` | Test without moving files |
| `--generate-testable-index` | Generate project-level image index |

### PDF Conversion

```bash
# Standard (pymupdf4llm + pdfplumber cross-validation)
python3 run-pipeline.py paper.pdf -o paper.md -i paper_images/

# Force MinerU for complex layouts
python3 run-pipeline.py scanned-doc.pdf --force-extractor mineru -o output.md

# Skip cross-validation for faster processing
python3 run-pipeline.py simple.pdf -o simple.md --skip-xval
```

**Extractor selection logic:**

| Document Type | Default Extractor | Fallback Chain |
|---|---|---|
| Digital PDF (>50 chars/page avg) | pymupdf4llm | markitdown &rarr; calibre |
| Scanned PDF (<50 chars/page avg) | tesseract | mineru &rarr; zerox |
| Complex PDF (>40% cross-val failures) | Auto-switches to MinerU | — |

### DOCX Conversion

```bash
# Standard (pandoc for text, python-docx for images)
python3 run-pipeline.py report.docx -o report.md

# With organized output
python3 run-pipeline.py report.docx --target-dir ./reports/
```

### PPTX Conversion

```bash
# Standard (python-pptx with recursive group shape extraction)
python3 run-pipeline.py deck.pptx -o deck.md

# Charts and SmartArt are rendered via LibreOffice when available
python3 run-pipeline.py charts.pptx --target-dir ./presentations/
```

### XLSX Conversion

XLSX files use a lightweight text-only path (no image pipeline):

```bash
# Via markitdown (recommended)
pip install markitdown
markitdown spreadsheet.xlsx > spreadsheet.md
```

### Organized Output (`--target-dir`)

When you specify `--target-dir`, the pipeline organizes all output:

```
target-dir/
  paper.md                    # Converted markdown
  paper_images/               # Extracted images
  paper_manifest.json         # Image manifest with metadata
  paper_image-index.md        # Image classification index
  _originals/                 # Source files moved here
    paper.pdf
  PIPELINE-REPORT.md          # Visual conversion report
  ISSUE-LOG.md                # Tracked issues (appended per conversion)
```

---

## Pipeline Steps

The full pipeline runs these steps in sequence:

| Step | Name | Tool | Description |
|---|---|---|---|
| **0** | Extractor Router | Python | Detect format, measure text density, select extractor |
| **1** | Text + Image Extraction | Python | Run selected extractor (pymupdf4llm, convert-office, etc.) |
| **1b** | Cross-Validation | Python | Compare extraction against pdfplumber (PDF only) |
| **1c** | Early Image Index | Python | Pre-QC image index for MinerU output |
| **2** | Structural QC | Python | **GATE** — must PASS before proceeding |
| **3** | Image Analysis Prep | Python | Persona activation matrix, analysis manifest |
| **4** | IMAGE NOTEs | Claude | Multi-expert image descriptions (8 personas) |
| **5** | Content Fidelity QC | Claude | Verify no text was lost in conversion |
| **6a** | Number Extraction | Python | Extract numerical data (PDF only) |
| **6c** | Image Index | Python | Per-image SUB/DEC classification with 8 heuristics |
| **7-13** | File Organization | Python | Move, rename, registry update, visual report |

Steps 0-3 and 6 run automatically. Steps 4-5 require Claude Code (Tier 2).

---

## Image Classification

Each extracted image passes through an 8-heuristic classification chain:

1. **Blank detection** — 3-tier: file size (<2KB), pixel statistics, near-black analysis
2. **Dimension check** — Minimum size thresholds
3. **Aspect ratio** — Extreme ratios suggest decorative elements (banners, rules)
4. **Journal branding** — Small logos, publisher marks
5. **Color block detection** — Solid/near-solid color fills
6. **Low-density badge** — Small images with minimal visual information
7. **Page position heuristics** — Header/footer regions
8. **Vector content detection** — pymupdf `get_drawings()` count + area analysis

Images classified as **substantive (SUB)** proceed to Tier 2 analysis.
Images classified as **decorative (DEC)** are skipped, saving LLM tokens.

### Persona Activation Matrix

For substantive images, the pipeline maps each image type to relevant expert personas:

| Image Type | Always Active | Conditionally Active |
|---|---|---|
| Kaplan-Meier | Statistician, Viz Critic | Clinical Trialist, Epidemiologist, Health Economist |
| Forest Plot | Statistician, Viz Critic | Clinical Trialist, Regulatory Analyst |
| Tornado Diagram | Health Economist, Statistician, Viz Critic | Regulatory Analyst |
| Decision Tree | Model Architect, Health Economist | Clinical Trialist, Regulatory Analyst |
| Flow Chart | Viz Critic | Clinical Trialist (CONSORT), Regulatory (PRISMA), Model Architect |
| Scatter Plot | Statistician, Viz Critic | Health Economist (CE plane), Epidemiologist |

The full matrix covers 24+ image types across 8 personas. The `prepare-image-analysis.py`
script generates an `analysis-manifest.json` with per-image persona assignments, template
skeletons, and section context.

---

## Structural QC Engine

`qc-structural.py` runs automated quality checks and acts as a **gate** — the pipeline
stops if QC fails.

### Checks Performed

- **YAML header validation** — Required fields: `source_file`, `conversion_date`, `conversion_tool`, `fidelity_standard`, `document_type`
- **Section/heading count** — Detects missing or collapsed sections
- **Table column consistency** — Flags tables with inconsistent column counts
- **Table collapse detection** — Detects multi-column tables collapsed into fewer cells (numeric density heuristic)
- **Reference numbering** — Validates `[1]`-`[N]` sequential references
- **Encoding errors** — Catches mojibake and broken Unicode
- **Image index completeness** — Cross-references manifest against extracted files
- **Manifest consistency** — Validates manifest JSON against image index table
- **Markdown syntax** — Checks for common formatting errors

### Exit Codes

| Code | Meaning | Pipeline Action |
|---|---|---|
| 0 | PASS | Continue to next step |
| 1 | FAIL | Pipeline stops — fix required |
| 2 | WARN | Fix and rerun (do not proceed on WARN) |

---

## Claude Code Integration

### Hook: Enforce MD-First Reading

The included hook intercepts `Read` tool calls in Claude Code and redirects
PDF/DOCX/PPTX reads to their converted Markdown equivalents.

**Setup:**

1. Copy the hook script:
```bash
cp hooks/enforce-pdf-conversion.sh ~/.claude/hooks/
chmod +x ~/.claude/hooks/enforce-pdf-conversion.sh
```

2. Register in `~/.claude/settings.json`:
```json
{
  "hooks": {
    "PreToolUse": [
      {
        "matcher": "Read",
        "hooks": [
          {
            "type": "command",
            "command": "~/.claude/hooks/enforce-pdf-conversion.sh"
          }
        ]
      }
    ]
  }
}
```

**How it works:**

1. Hook intercepts every `Read` tool call
2. If the file is a PDF/DOCX/PPTX:
   - Computes SHA-256 hash
   - Looks up the hash in the conversion registry
   - If found: redirects to the registered `.md` file
   - If not found: checks for a co-located `.md` (same directory, same name)
   - If no `.md` exists: **blocks the read** and prints the conversion command
3. All other file types pass through unchanged
4. Every interception is logged to `~/.claude/pipeline/hook-interceptions.log`

### Skill: Full Pipeline Orchestration

The included `SKILL.md` defines a Claude Code skill that orchestrates the
complete pipeline with step-by-step instructions for both tiers.

**Setup:**

```bash
cp skill/SKILL.md ~/.claude/skills/convert-documents/SKILL.md
```

The skill provides:
- Quick-start commands for each format
- Step-by-step orchestration instructions
- Expert persona reference table
- QC loop enforcement (fix and rerun until zero issues)
- Image analysis prompt templates

---

## Configuration

### Conversion Registry

The pipeline maintains a JSON registry at `~/.claude/pipeline/conversion_registry.json`.
Each entry records:

```json
{
  "sha256": "a1b2c3...",
  "source_file": "/path/to/original.pdf",
  "output_md": "/path/to/converted.md",
  "pipeline_version": "3.2.0",
  "extractor": "pymupdf4llm",
  "conversion_date": "2025-01-15T10:30:00Z",
  "pages": 47,
  "image_index_path": "/path/to/image-index.md",
  "total_images_detected": 30,
  "substantive_images": 22,
  "has_testable_images": true
}
```

The registry enables:
- **Deduplication** — Same file (by hash) is never converted twice
- **Hook lookup** — The Claude Code hook finds converted `.md` by hash
- **Audit trail** — Full provenance for every conversion

### Image Index Overrides

For cases where automatic classification is wrong, create an
`image-index-overrides.json` alongside the image index:

```json
{
  "page_5_img_3.png": {
    "classification": "SUB",
    "reason": "Manual override: contains relevant diagram"
  }
}
```

The pipeline applies overrides during image index generation (Step 6c).

### Output Metadata

Every converted Markdown file includes a YAML frontmatter header:

```yaml
---
source_file: paper.pdf
source_format: pdf
conversion_date: "2025-01-15T10:30:00Z"
conversion_tool: pymupdf4llm
pipeline_version: "3.2.0"
fidelity_standard: zero_missing_text
document_type: academic_paper
pages: 47
domain: health_economics
---
```

---

## Component Reference

| File | Lines | Description |
|---|---|---|
| `scripts/run-pipeline.py` | 7,749 | Unified pipeline router, image classification, file organization, registry management |
| `scripts/convert-paper.py` | 1,329 | PDF text/image extraction via pymupdf4llm, multi-panel splitting, sparse page rendering |
| `scripts/convert-office.py` | 3,122 | DOCX/PPTX conversion, recursive shape extraction, chart rendering, PUA Unicode mapping |
| `scripts/qc-structural.py` | 1,211 | Structural QC engine: YAML validation, table collapse detection, encoding checks |
| `scripts/prepare-image-analysis.py` | 634 | Persona activation matrix, analysis manifest generation, template skeletons |
| `scripts/convert-mineru.py` | 228 | MinerU fallback wrapper for complex/scanned PDFs (CPU-only) |
| `scripts/style-docx-tables.py` | 262 | Professional DOCX styling for pandoc output (table colors, borders, code blocks) |
| `hooks/enforce-pdf-conversion.sh` | 276 | Claude Code PreToolUse hook: intercepts PDF/Office reads, redirects to Markdown |
| `skill/SKILL.md` | 1,354 | Claude Code skill definition: full pipeline orchestration with QC loops |

**Total: ~16,165 lines**

---

## Dependencies

### Required

| Package | Purpose |
|---|---|
| [PyMuPDF](https://pymupdf.readthedocs.io/) (`fitz`) | PDF parsing, image extraction, vector detection |
| [pymupdf4llm](https://github.com/pymupdf/pymupdf4llm) | LLM-optimized Markdown extraction from PDF |
| [pdfplumber](https://github.com/jsvine/pdfplumber) | Cross-validation of PDF extraction |
| [python-docx](https://python-docx.readthedocs.io/) | DOCX image extraction and styling |
| [python-pptx](https://python-pptx.readthedocs.io/) | PPTX text and image extraction |
| [Pillow](https://pillow.readthedocs.io/) | Image processing, blank detection, format conversion |
| [NumPy](https://numpy.org/) | Pixel-level image analysis (near-black detection) |
| [Pandoc](https://pandoc.org/) | DOCX text extraction to Markdown |
| [jq](https://stedolan.github.io/jq/) | JSON processing in the hook script |

### Optional

| Package | Purpose |
|---|---|
| [LibreOffice](https://www.libreoffice.org/) | Chart/SmartArt rendering (PPTX) |
| [MinerU](https://github.com/opendatalab/MinerU) | Complex/scanned PDF fallback extractor |
| [Tesseract](https://github.com/tesseract-ocr/tesseract) | OCR for scanned documents |
| [MarkItDown](https://github.com/microsoft/markitdown) | XLSX and fallback PDF conversion |

### Python Version

Python 3.8+ is required. The codebase uses `dataclasses`, `typing.Literal`,
and `pathlib` features available from Python 3.8 onward.

---

## Examples

### Convert an Academic Paper

```bash
# Full pipeline with organized output
python3 scripts/run-pipeline.py \
  ~/papers/smith-2024-cost-effectiveness.pdf \
  --target-dir ~/converted/smith-2024/

# Output structure:
# ~/converted/smith-2024/
#   smith-2024-cost-effectiveness.md
#   smith-2024-cost-effectiveness_images/
#   smith-2024-cost-effectiveness_manifest.json
#   smith-2024-cost-effectiveness_image-index.md
#   _originals/smith-2024-cost-effectiveness.pdf
#   PIPELINE-REPORT.md
```

### Convert a Slide Deck with Charts

```bash
# Charts are rendered via LibreOffice at 300 DPI
python3 scripts/run-pipeline.py \
  ~/presentations/quarterly-review.pptx \
  --target-dir ~/converted/quarterly/

# SmartArt and charts appear as high-resolution PNG images
# in the _images/ directory with type_guess="chart" or "diagram"
```

### Batch Conversion

```bash
# Convert all PDFs in a directory
for f in ~/papers/*.pdf; do
  python3 scripts/run-pipeline.py "$f" \
    --target-dir ~/converted/ \
    --skip-xval
done

# Generate project-level image index
python3 scripts/run-pipeline.py --generate-testable-index ~/converted/
```

### Dry Run (Preview Without Moving Files)

```bash
python3 scripts/run-pipeline.py paper.pdf \
  --target-dir ~/converted/ \
  --dry-run
```

---

## Troubleshooting

### Common Issues

| Issue | Cause | Fix |
|---|---|---|
| `FAIL: No YAML header block found` | Extractor produced malformed output | Check source file is valid; try `--force-extractor mineru` |
| Step 2 WARN: table collapse | Multi-column tables lost columns in conversion | QC inserts HTML WARNING comments; fix manually or re-extract |
| MinerU fallback triggered | >40% of pages failed cross-validation | Expected for complex layouts; MinerU handles these better |
| `ValueError: min() iterable argument is empty` | pymupdf4llm bug on certain table layouts | Fixed by disabling layout mode; should not recur |
| Hook blocks PDF read | No converted `.md` found | Run the pipeline first: `python3 run-pipeline.py <file>` |
| Near-black images not detected | Anti-aliased rendering creates subtle gradients | Pipeline uses 4-tier detection including pixel-percentage pass |

### Exit Codes

| Script | Code | Meaning |
|---|---|---|
| `run-pipeline.py` | 0 | Success |
| `run-pipeline.py` | 1 | General failure |
| `run-pipeline.py` | 3 | Extractor crash (pymupdf4llm) |
| `qc-structural.py` | 0 | QC PASS |
| `qc-structural.py` | 1 | QC FAIL |
| `qc-structural.py` | 2 | QC WARN |
| `convert-mineru.py` | 1 | MinerU not installed |
| `convert-mineru.py` | 2 | Conversion failed |

---

## Design Decisions

**Why not just use `markitdown`?**
MarkItDown is excellent for simple documents but loses table structure,
heading hierarchy, and images in complex PDFs. This pipeline uses
pymupdf4llm for superior table and multi-column support, with pdfplumber
cross-validation to catch extraction errors.

**Why a 2-tier architecture?**
LLM tokens are expensive. The Python tier handles everything that can be
done deterministically (text extraction, image classification, QC) at zero
token cost. The LLM tier is reserved for tasks that genuinely require
visual understanding (image descriptions) or natural language judgement
(content fidelity verification).

**Why 8 expert personas?**
A single "describe this image" prompt produces generic descriptions.
Domain-specific personas (e.g., a Statistician analyzing a Kaplan-Meier
curve) produce descriptions that capture methodologically relevant details
like confidence intervals, at-risk tables, and crossing hazard curves.

**Why SHA-256 everywhere?**
File names change. File contents don't. Hash-based deduplication and registry
lookup means the pipeline never re-converts a document it has already processed,
even if the file is moved, renamed, or copied to a different directory.

---

## Contributing

Contributions are welcome. Please:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/my-feature`)
3. Run the QC checks on any modified scripts
4. Submit a pull request with a clear description of changes

### Development Setup

```bash
git clone https://github.com/orangefineblue/doc2md.git
cd doc2md
pip install -r requirements.txt  # when available

# Run structural QC on a test conversion
python3 scripts/run-pipeline.py tests/fixtures/sample.pdf -o /tmp/test.md
python3 scripts/qc-structural.py /tmp/test.md --verbose
```

### Reporting Issues

When reporting a bug, please include:
- The source file format (PDF/DOCX/PPTX)
- The extractor used (check pipeline output)
- The full error message or QC failure output
- Python version (`python3 --version`)

---

## License

MIT License. See [LICENSE](LICENSE) for details.

---

## Related Projects

- **[claude-code-orchestration-protocol](https://github.com/orangefineblue/claude-code-orchestration-protocol)** — A zero-read orchestrator protocol for Claude Code that manages context window usage, delegates work to sub-agents, and runs QC loops until zero issues remain. Designed to work alongside doc2md for complex multi-document workflows where context rot is a concern.

---

## Acknowledgements

Built for use with [Claude Code](https://docs.anthropic.com/en/docs/claude-code)
by Anthropic. Uses [PyMuPDF](https://pymupdf.readthedocs.io/),
[pdfplumber](https://github.com/jsvine/pdfplumber),
[MinerU](https://github.com/opendatalab/MinerU), and
[Pandoc](https://pandoc.org/) for document processing.

---

<div align="center">

**doc2md** is designed for researchers, analysts, and anyone who needs
high-fidelity document conversion in LLM-powered workflows.

[Report a Bug](https://github.com/orangefineblue/doc2md/issues) |
[Request a Feature](https://github.com/orangefineblue/doc2md/issues)

</div>
