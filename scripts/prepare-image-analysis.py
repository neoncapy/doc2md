#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Prepare Image Analysis - Persona Activation Matrix (NEW script).

Pre-computes which expert personas should analyze each image.
Generates analysis-manifest.json for Claude's IMAGE NOTE generation.

This runs at ZERO token cost — all logic is Python heuristics.

Inputs:
  - markdown file
  - image-manifest.json (auto-detected from MD)
  - context-summary.json (auto-detected)

Outputs:
  - analysis-manifest.json in images directory

The analysis manifest contains:
  - Per-image activation matrix (which personas analyze which images)
  - Template skeletons (pre-filled IMAGE NOTE fields)
  - Section context snippets
  - Persona definitions

Usage:
    python3 prepare-image-analysis.py <converted.md>
"""

import argparse
import json
import re
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional

# The full persona activation matrix from discovery-personas-design.md
ACTIVATION_MATRIX = {
    "kaplan-meier": {
        "always": ["statistician", "visualization_critic"],
        "conditional": {
            "clinical_trialist": {"context_keywords": ["trial", "arm", "treatment", "group", "randomized", "RCT"]},
            "epidemiologist": {"context_keywords": ["population", "registry", "cohort", "incidence", "prevalence"]},
            "health_economist": {"context_keywords": ["extrapolat", "survival model", "partitioned", "QALY", "cost"]},
        }
    },
    "forest-plot": {
        "always": ["statistician", "visualization_critic"],
        "conditional": {
            "clinical_trialist": {"context_keywords": ["subgroup", "trial", "ITT", "per-protocol", "arm"]},
            "regulatory_analyst": {"context_keywords": ["HTA", "systematic review", "submission", "NICE", "DMP", "PRISMA"]},
        }
    },
    "tornado-diagram": {
        "always": ["health_economist", "statistician", "visualization_critic"],
        "conditional": {
            "regulatory_analyst": {"context_keywords": ["HTA", "submission", "sensitivity", "DSA"]},
        }
    },
    "scatter": {
        "always": ["statistician", "visualization_critic"],
        "conditional": {
            "health_economist": {"context_keywords": ["cost", "ICER", "QALY", "WTP", "incremental", "CE plane"]},
            "epidemiologist": {"context_keywords": ["population", "registry", "cohort", "correlation"]},
            "regulatory_analyst": {"context_keywords": ["WTP", "threshold", "CE plane", "HTA"]},
        }
    },
    "decision-tree": {
        "always": ["model_architect", "health_economist"],
        "conditional": {
            "clinical_trialist": {"context_keywords": ["treatment", "pathway", "clinical", "arm", "strategy"]},
            "regulatory_analyst": {"context_keywords": ["HTA", "submission", "model", "DMP", "NICE"]},
        }
    },
    "flow-chart": {
        "always": ["visualization_critic"],
        "conditional": {
            "clinical_trialist": {"context_keywords": ["CONSORT", "patient", "randomized", "enrollment", "trial"]},
            "regulatory_analyst": {"context_keywords": ["PRISMA", "systematic", "review", "regulatory", "process"]},
            "model_architect": {"context_keywords": ["model", "algorithm", "logic", "decision", "analysis"]},
            "health_economist": {"context_keywords": ["economic", "model", "cost", "analysis"]},
        }
    },
    "schematic": {
        "always": ["visualization_critic"],
        "conditional": {
            "model_architect": {"context_keywords": ["model", "Markov", "state", "transition", "structure"]},
            "health_economist": {"context_keywords": ["economic", "model", "cost"]},
            "clinical_trialist": {"context_keywords": ["clinical", "pathway", "treatment", "disease"]},
        }
    },
    "line-chart": {
        "always": ["visualization_critic"],
        "conditional": {
            "statistician": {"context_keywords": ["model fit", "statistical", "trend", "regression"]},
            "epidemiologist": {"context_keywords": ["time trend", "survival", "incidence", "mortality", "temporal"]},
            "health_economist": {"context_keywords": ["cost", "projection", "CEAC", "acceptability", "budget"]},
        }
    },
    "bar-chart": {
        "always": ["visualization_critic"],
        "conditional": {
            "epidemiologist": {"context_keywords": ["disease burden", "prevalence", "incidence", "DALY", "population"]},
            "health_economist": {"context_keywords": ["cost", "breakdown", "budget impact", "resource use"]},
            "clinical_trialist": {"context_keywords": ["adverse event", "response rate", "outcome", "trial"]},
        }
    },
    "table-image": {
        "always": ["visualization_critic"],
        "conditional": {
            "health_economist": {"context_keywords": ["cost", "QALY", "ICER", "economic", "utility"]},
            "clinical_trialist": {"context_keywords": ["patient characteristic", "endpoint", "baseline", "outcome"]},
            "statistician": {"context_keywords": ["statistical", "CI", "p-value", "result", "analysis"]},
        }
    },
    "histogram": {
        "always": ["statistician", "visualization_critic"],
        "conditional": {}
    },
    "box-plot": {
        "always": ["statistician", "visualization_critic"],
        "conditional": {
            "clinical_trialist": {"context_keywords": ["clinical", "outcome", "treatment", "arm"]},
        }
    },
    "heatmap": {
        "always": ["statistician", "visualization_critic"],
        "conditional": {}
    },
    "pie-chart": {
        "always": ["visualization_critic"],
        "conditional": {
            "epidemiologist": {"context_keywords": ["disease", "distribution", "population", "proportion"]},
        }
    },
    "funnel-plot": {
        "always": ["statistician", "visualization_critic"],
        "conditional": {}
    },
    "log_cumulative_hazard_plot": {
        "always": ["statistician"],
        "conditional": {
            "clinical_trialist": {"context_keywords": ["trial", "treatment", "arm", "proportional hazard"]},
        }
    },
    "network-diagram": {
        "always": ["statistician", "visualization_critic"],
        "conditional": {}
    },
    "photograph": {
        "always": [],
        "conditional": {
            "epidemiologist": {"context_keywords": ["disease", "manifestation", "clinical"]},
        }
    },
    "other": {
        "always": ["visualization_critic"],
        "conditional": {}
    },
    "process_flowchart": {
        "always": ["visualization_critic"],
        "conditional": {
            "regulatory_analyst": {"context_keywords": ["regulatory", "compliance", "process", "procedure"]},
        }
    },
    "decorative-background": {
        "always": [],
        "conditional": {}
    },
    # Generic "chart" and "diagram" types from convert-office.py chart/SmartArt
    # renders. These are full-slide renders where specific chart subtype is unknown.
    "chart": {
        "always": ["statistician", "visualization_critic", "health_economist", "model_architect"],
        "conditional": {
            "clinical_trialist": {"context_keywords": ["trial", "arm", "treatment", "endpoint"]},
            "regulatory_analyst": {"context_keywords": ["HTA", "submission", "NICE"]},
        }
    },
    "diagram": {
        "always": ["visualization_critic", "model_architect"],
        "conditional": {
            "health_economist": {"context_keywords": ["economic", "model", "cost", "QALY"]},
            "clinical_trialist": {"context_keywords": ["clinical", "pathway", "treatment"]},
            "regulatory_analyst": {"context_keywords": ["HTA", "process", "regulatory"]},
        }
    },
}

# Persona definitions for the analysis manifest
PERSONA_DEFINITIONS = {
    "statistician": "Senior biostatistician: statistical methodology, CIs, p-values, assumptions",
    "health_economist": "Senior health economist: ICERs, QALYs, WTP, model parameters",
    "clinical_trialist": "Senior clinical researcher: trial design, CONSORT, clinical plausibility",
    "model_architect": "Senior decision modeler: Markov, microsimulation, structural validity",
    "visualization_critic": "Data visualization expert: honest representation, accessibility",
    "regulatory_analyst": "Senior HTA analyst: NICE/DMP/NoMA compliance, reference case",
    "epidemiologist": "Senior epidemiologist: registry data, population validity, disease burden"
}


def load_manifest(md_path: Path, explicit_manifest: Optional[Path] = None) -> dict:
    """Load image-manifest.json from explicit path or auto-detected location."""
    # BUG-2 FIX: use explicit --manifest path when provided
    if explicit_manifest is not None:
        if not explicit_manifest.exists():
            print(f"ERROR: manifest file not found: {explicit_manifest}")
            sys.exit(1)
        with open(explicit_manifest, 'r', encoding='utf-8') as f:
            return json.load(f)

    # Auto-detect when --manifest is not provided
    manifest_path = find_manifest_path(md_path)
    if not manifest_path or not manifest_path.exists():
        print(f"ERROR: image-manifest.json not found")
        print(f"  Expected: {manifest_path}")
        sys.exit(1)

    with open(manifest_path, 'r', encoding='utf-8') as f:
        return json.load(f)


def load_context_summary(md_path: Path) -> Optional[dict]:
    """Load context-summary.json if it exists."""
    summary_path = md_path.parent / "context-summary.json"
    if not summary_path.exists():
        print("WARNING: context-summary.json not found (using defaults)")
        return None

    with open(summary_path, 'r', encoding='utf-8') as f:
        return json.load(f)


def extract_title_from_yaml_frontmatter(md_path: Path) -> Optional[str]:
    """Extract the title field from YAML frontmatter of a converted .md file.

    Frontmatter is the block between the first --- and the second --- lines.
    Returns the stripped title string, or None if not found.

    QC-5: Used as a fallback when context-summary.json is missing or has a
    wrong/empty title. The YAML frontmatter is always written by convert-office.py
    and convert-paper.py, so it is the most reliable title source.
    """
    try:
        content = md_path.read_text(encoding='utf-8')
        # Must start with --- on the first line
        if not content.startswith('---'):
            return None
        # Find the closing ---
        end = content.find('\n---', 3)
        if end == -1:
            return None
        frontmatter = content[3:end]
        # Parse title: field (simple regex, avoids yaml dep)
        # Handle multi-line YAML titles by joining continuation lines
        # that follow the title: key (indented or part of a quoted string).
        fm_lines = frontmatter.splitlines()
        for i, line in enumerate(fm_lines):
            m = re.match(r'^title:\s*"?([^"]*)"?\s*$', line.strip())
            if m:
                title = m.group(1).strip().strip('"').strip("'")
                # Check for continuation lines (YAML folded/literal or
                # quoted string with embedded newline)
                j = i + 1
                while j < len(fm_lines):
                    cont = fm_lines[j]
                    # Stop at next key (unindented key: value)
                    if re.match(r'^[a-zA-Z_]', cont):
                        break
                    # Continuation line (indented or empty)
                    title += " " + cont.strip().strip('"').strip("'")
                    j += 1
                # Normalize whitespace (collapse newlines, extra spaces)
                title = " ".join(title.split())
                if title:
                    return title
        return None
    except Exception:
        return None


def _extract_domain_from_yaml_frontmatter(md_path: Path) -> Optional[str]:
    """Extract document_domain from YAML frontmatter of a converted .md file.

    S21/RC2: Used as a fallback when context-summary.json is missing or has
    domain "general".  The YAML frontmatter is written by run-pipeline.py
    (F14) after domain detection, so it is a reliable domain source.
    Returns the domain string, or None if not found.
    """
    try:
        content = md_path.read_text(encoding='utf-8')
        if not content.startswith('---'):
            return None
        end = content.find('\n---', 3)
        if end == -1:
            return None
        frontmatter = content[3:end]
        for line in frontmatter.splitlines():
            m = re.match(r'^document_domain:\s*(\S+)', line.strip())
            if m:
                domain = m.group(1).strip().strip('"').strip("'")
                if domain and domain.lower() != "general":
                    return domain
        return None
    except Exception:
        return None


def find_manifest_path(md_path: Path) -> Path:
    """Find image-manifest.json path from markdown file.

    Handles two formats:
    - PDF (convert-paper.py): image links are 'images/{subdir}/fig1.png'
    - Office (convert-office.py): image links are bare filenames like 's02-img01.png'
      with images in '{basename}_images/' and manifest as '{basename}_manifest.json'
    """
    content = md_path.read_text(encoding='utf-8')

    # --- PDF pattern: links of the form (images/subdir/filename.png) ---
    img_links = re.findall(r'\[([^\]]+)\]\((images/[^)]+)\)', content)
    if img_links:
        first_link = img_links[0][1]
        images_dir_match = re.match(r'(images/[^/]+)', first_link)
        if images_dir_match:
            images_dir = md_path.parent / images_dir_match.group(1)
            return images_dir / "image-manifest.json"

    # --- Office pattern: check for {basename}_manifest.json alongside the .md ---
    # BUG-3 FIX: convert-office.py writes bare filenames (e.g. s02-img01.png)
    # and places the manifest as {basename}_manifest.json next to the .md file.
    office_manifest = md_path.parent / f"{md_path.stem}_manifest.json"
    if office_manifest.exists():
        return office_manifest

    # Also try: bare filename links (![...](s02-img01.png)) hint at office format,
    # look for any *_manifest.json in the same directory.
    bare_links = re.findall(r'\[([^\]]+)\]\(([^/)][^)]*\.(?:png|jpg|jpeg|gif|bmp|tiff))\)', content)
    if bare_links:
        # Search for *_manifest.json in the same dir
        for candidate in sorted(md_path.parent.glob("*_manifest.json")):
            return candidate

    # Default fallback (PDF canonical path)
    short_name = md_path.stem
    default_dir = md_path.parent / "images" / short_name
    return default_dir / "image-manifest.json"


def check_conditional_triggers(img_type, image_entry, context_domain):
    # type: (str, dict, str) -> list
    """
    Check if conditional personas should be activated.
    Returns list of activated conditional personas.
    """
    activated = []
    conditionals = ACTIVATION_MATRIX.get(img_type, {}).get("conditional", {})

    # Build searchable text from image metadata
    searchable_text = " ".join([
        image_entry.get("detected_caption") or "",
        image_entry.get("nearby_text") or "",
        (image_entry.get("section_context") or {}).get("heading", ""),
        context_domain,
    ]).lower()

    for persona, rules in conditionals.items():
        keywords = rules.get("context_keywords", [])
        if any(kw.lower() in searchable_text for kw in keywords):
            activated.append(persona)

    return activated


def compute_activation_for_image(image_entry: dict, context_summary: Optional[dict]) -> dict:
    """
    Compute which personas should analyze this image.
    Returns dict with activated_personas, conditional_personas, and template_skeleton.
    """
    img_type = image_entry.get("type_guess") or "other"
    context_domain = (context_summary or {}).get("document_domain", "general")

    # Get always-active personas
    always_active = ACTIVATION_MATRIX.get(img_type, {}).get("always", [])

    # Check conditional personas
    conditional_active = check_conditional_triggers(img_type, image_entry, context_domain)

    # Remove duplicates
    all_active = list(dict.fromkeys(always_active + conditional_active))

    # Conditional personas that were NOT activated
    all_conditionals = list(ACTIVATION_MATRIX.get(img_type, {}).get("conditional", {}).keys())
    conditional_not_active = [p for p in all_conditionals if p not in all_active]

    # Build template skeleton (fields Python can fill)
    template_skeleton = {
        "type": img_type,
        "file": image_entry.get("file_path", f"images/{image_entry['filename']}"),  # Use file_path from manifest
        "data_density": None,  # Claude fills
        "context": None,  # Claude fills
        "readable_at": None,  # Claude fills
    }

    return {
        "activated_personas": all_active,
        "conditional_personas": {
            p: f"activate if {', '.join(ACTIVATION_MATRIX.get(img_type, {}).get('conditional', {}).get(p, {}).get('context_keywords', []))}"
            for p in conditional_not_active
        },
        "template_skeleton": template_skeleton,
    }


def _relative_file_path(img: dict, manifest: dict, md_path: Path) -> str:
    """
    Compute a relative file_path for the analysis manifest entry.

    For PDF (convert-paper.py): images are in images/{short_name}/ relative to
    md_path.parent, so the path is like 'images/short-name/fig1.png'.

    For office (convert-office.py): images are in {basename}_images/ relative to
    md_path.parent, so the path is like 'presentation_images/s01-img01.png'.

    Priority:
    1. img["file_path"] if present and absolute — compute relative to md_path.parent
    2. images_dir from manifest + img["filename"] — compute relative to md_path.parent
    3. Legacy fallback: 'images/{filename}' (original hardcoded behavior for PDF)
    """
    import os

    filename = img.get("filename", "")
    md_dir = md_path.parent

    # Option 1: use absolute file_path from manifest entry
    abs_path_str = img.get("file_path")
    if abs_path_str:
        try:
            rel = os.path.relpath(abs_path_str, str(md_dir))
            return rel
        except ValueError:
            pass  # Windows cross-drive; fall through

    # Option 2: use images_dir from manifest top-level
    images_dir_str = manifest.get("images_dir")
    if images_dir_str and filename:
        try:
            abs_img = Path(images_dir_str) / filename
            rel = os.path.relpath(str(abs_img), str(md_dir))
            return rel
        except ValueError:
            pass

    # Option 3: legacy fallback (PDF canonical path)
    return f"images/{filename}"


def generate_analysis_manifest(md_path: Path, manifest: dict, context_summary: Optional[dict]) -> dict:
    """
    Generate the full analysis-manifest.json.

    QC-4: Skips images marked is_duplicate=True, blank=True, or decorative=True.
           These are tracked but not sent to Opus vision analysis.
    QC-5: document_title falls back to YAML frontmatter title when
           context-summary.json is missing or has an empty/unknown title.
    """
    analysis_images = []
    total_persona_analyses = 0
    skipped_duplicates = 0
    skipped_blanks = 0
    skipped_decoratives = 0
    all_images = manifest.get("images", [])
    # Per-image classification from image index (populated by run-pipeline.py)
    _per_img_class = manifest.get("per_image_classification", {})

    for img in all_images:
        # QC-4: Skip byte-identical duplicate images
        if img.get("is_duplicate"):
            skipped_duplicates += 1
            continue

        # QC-3/QC-4: Skip blank images (WMF conversion failures)
        if img.get("blank"):
            skipped_blanks += 1
            continue

        # Skip decorative images (e.g. slide backgrounds, borders)
        if img.get("decorative"):
            skipped_decoratives += 1
            continue

        # Skip images classified as DEC by image index (per-image)
        _img_fname = img.get("filename", "")
        if _img_fname in _per_img_class:
            if _per_img_class[_img_fname].get("classification") == "DEC":
                skipped_decoratives += 1
                continue

        activation = compute_activation_for_image(img, context_summary)
        total_persona_analyses += len(activation["activated_personas"])

        # Compute correct relative path (works for both PDF and office formats)
        file_path_rel = _relative_file_path(img, manifest, md_path)

        analysis_images.append({
            "figure_num": img["figure_num"],
            "filename": img["filename"],
            "file_path": file_path_rel,
            "page": img.get("page"),
            "dimensions": {"width": img["width"], "height": img["height"]},
            "detected_caption": img.get("detected_caption"),
            "type_guess": img.get("type_guess"),
            "section_context": img.get("section_context"),
            "activated_personas": activation["activated_personas"],
            "conditional_personas": activation["conditional_personas"],
            "template_skeleton": activation["template_skeleton"],
        })

    if skipped_duplicates > 0 or skipped_blanks > 0 or skipped_decoratives > 0:
        print(
            f"  Skipping {skipped_duplicates} duplicate image(s), "
            f"{skipped_blanks} blank image(s), and "
            f"{skipped_decoratives} decorative image(s) from vision analysis"
        )

    # Compute persona frequency
    persona_counts = {}
    for img in analysis_images:
        for persona in img["activated_personas"]:
            persona_counts[persona] = persona_counts.get(persona, 0) + 1

    most_common_persona = None
    if persona_counts:
        most_common_persona = max(persona_counts.items(), key=lambda x: x[1])

    # QC-5: document_title resolution — 3-tier priority:
    # 1. context-summary.json title (if present and non-empty/non-"Unknown")
    # 2. YAML frontmatter title from the converted .md file
    # 3. Fallback: "Unknown"
    ctx_title = (context_summary or {}).get("title", "").strip()
    if ctx_title and ctx_title.lower() not in ("unknown", ""):
        document_title = ctx_title
    else:
        yaml_title = extract_title_from_yaml_frontmatter(md_path)
        if yaml_title and yaml_title.lower() not in ("unknown", ""):
            document_title = yaml_title
        else:
            document_title = ctx_title or "Unknown"

    # S21/RC2: document_domain resolution — 3-tier priority:
    # 1. context-summary.json domain (if present and non-"general")
    # 2. YAML frontmatter document_domain (written by F14 in run-pipeline.py)
    # 3. Fallback: "general"
    ctx_domain = (context_summary or {}).get("document_domain", "general")
    if ctx_domain and ctx_domain.lower() != "general":
        document_domain = ctx_domain
    else:
        yaml_domain = _extract_domain_from_yaml_frontmatter(md_path)
        if yaml_domain:
            document_domain = yaml_domain
        else:
            document_domain = ctx_domain or "general"

    return {
        "md_file": str(md_path),
        "images_dir": str(manifest["images_dir"]),
        "document_domain": document_domain,
        "document_title": document_title,
        "total_images": len(all_images),
        "images_for_analysis": len(analysis_images),
        "skipped_duplicates": skipped_duplicates,
        "skipped_blanks": skipped_blanks,
        "skipped_decoratives": skipped_decoratives,
        "analysis_date": datetime.now().isoformat(),
        "images": analysis_images,
        "persona_definitions": PERSONA_DEFINITIONS,
        "activation_summary": {
            "total_persona_analyses": total_persona_analyses,
            "avg_personas_per_image": round(total_persona_analyses / max(1, len(analysis_images)), 1),
            "most_common_persona": f"{most_common_persona[0]} ({most_common_persona[1]}/{len(analysis_images)} images)" if most_common_persona else None,
            "persona_frequencies": persona_counts,
        }
    }


def main():
    parser = argparse.ArgumentParser(
        description="Prepare image analysis: compute persona activation matrix"
    )
    parser.add_argument("md_file", type=Path, help="Path to converted markdown file")
    parser.add_argument(
        "--manifest", type=Path, default=None,
        help="Path to image-manifest.json (auto-detected if omitted)"
    )
    parser.add_argument(
        "--context-summary", type=Path, default=None,
        help="Path to context-summary.json (auto-detected if omitted)"
    )

    args = parser.parse_args()

    if not args.md_file.exists():
        print(f"ERROR: File not found: {args.md_file}")
        sys.exit(1)

    print(f"Preparing image analysis for: {args.md_file}")

    # Load inputs
    # BUG-2 FIX: pass args.manifest explicitly so --manifest argument is honoured
    manifest = load_manifest(args.md_file, explicit_manifest=args.manifest)
    context_summary = load_context_summary(args.md_file)

    print(f"  Images: {len(manifest.get('images', []))}")
    if context_summary:
        print(f"  Domain: {context_summary.get('document_domain', 'general')}")

    # Generate analysis manifest
    analysis_manifest = generate_analysis_manifest(args.md_file, manifest, context_summary)

    # Write output
    images_dir = Path(manifest["images_dir"])
    output_path = images_dir / "analysis-manifest.json"
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(analysis_manifest, f, indent=2, ensure_ascii=False)

    print(f"\nDone. Analysis manifest written to: {output_path}")
    print(f"  Total persona-analyses: {analysis_manifest['activation_summary']['total_persona_analyses']}")
    print(f"  Avg personas/image: {analysis_manifest['activation_summary']['avg_personas_per_image']}")
    if analysis_manifest['activation_summary']['most_common_persona']:
        print(f"  Most common: {analysis_manifest['activation_summary']['most_common_persona']}")


if __name__ == "__main__":
    main()
