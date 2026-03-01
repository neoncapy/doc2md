#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MinerU Wrapper for PDF-to-Markdown Conversion.

Calls MinerU from its dedicated venv (~/envs/mineru/) to convert
scanned or complex PDFs to Markdown. Used as a fallback extractor
when Tesseract OCR is insufficient.

SAFETY: CPU-only mode enforced. NEVER use MPS/GPU on M4 Pro Mac
(causes kernel panics - see ~/.claude/projects MEMORY.md).

Usage:
    python3 convert-mineru.py <input.pdf> --output <output.md>

Exit codes:
    0 = success
    1 = MinerU not installed or venv missing
    2 = conversion failed
"""

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


MINERU_VENV = Path.home() / "envs" / "mineru"
MINERU_PYTHON = MINERU_VENV / "bin" / "python3"
MINERU_MAGIC_PDF = MINERU_VENV / "bin" / "magic-pdf"


def _warn(msg: str) -> None:
    """Log warning to stderr."""
    print(f"WARN [convert-mineru]: {msg}", file=sys.stderr)


def _fail(msg: str, exit_code: int = 2) -> None:
    """Log failure to stderr and exit."""
    print(f"FAIL [convert-mineru]: {msg}", file=sys.stderr)
    sys.exit(exit_code)


def check_mineru_installed() -> bool:
    """Verify MinerU venv and magic-pdf binary exist."""
    if not MINERU_VENV.exists():
        _warn(f"MinerU venv not found at {MINERU_VENV}")
        return False

    if not MINERU_PYTHON.exists():
        _warn(f"Python not found in MinerU venv at {MINERU_PYTHON}")
        return False

    # Check for magic-pdf CLI (MinerU's entry point)
    if not MINERU_MAGIC_PDF.exists():
        # Try finding it via the venv python
        result = subprocess.run(
            [str(MINERU_PYTHON), "-m", "magic_pdf", "--help"],
            capture_output=True, timeout=30
        )
        if result.returncode != 0:
            _warn("magic-pdf (MinerU CLI) not found in venv")
            return False

    return True


def convert_with_mineru(pdf_path: Path, output_md: Path) -> bool:
    """Run MinerU conversion. CPU-only mode enforced.

    MinerU outputs to a directory. We find the .md file and move it
    to the requested output path.

    Args:
        pdf_path: Path to input PDF.
        output_md: Path for output Markdown file.

    Returns:
        True if conversion succeeded, False otherwise.
    """
    # NOTE: TemporaryDirectory context manager guarantees cleanup on exit
    # (normal or exception). If shutil.copy2() fails after context exit,
    # the converted content is lost. This is acceptable: the caller
    # retries with the next extractor in the fallback chain.
    with tempfile.TemporaryDirectory(prefix="mineru-") as tmpdir:
        tmpdir_path = Path(tmpdir)

        # Build MinerU command
        # Force CPU via environment variables (NEVER use MPS/GPU)
        env = os.environ.copy()
        env["CUDA_VISIBLE_DEVICES"] = ""
        env["PYTORCH_MPS_DISABLE"] = "1"
        env["PYTORCH_ENABLE_MPS_FALLBACK"] = "0"

        # Try magic-pdf CLI first
        if MINERU_MAGIC_PDF.exists():
            cmd = [
                str(MINERU_MAGIC_PDF),
                "-p", str(pdf_path),
                "-o", str(tmpdir_path),
                "-m", "auto",  # auto-detect OCR vs text mode
            ]
        else:
            # Fallback: invoke via python -m
            # Module is "magic_pdf" (not "magic_pdf.cli"), matching
            # what check_mineru_installed() tests with --help
            cmd = [
                str(MINERU_PYTHON),
                "-m", "magic_pdf",
                "-p", str(pdf_path),
                "-o", str(tmpdir_path),
                "-m", "auto",
            ]

        print(f"  Running MinerU (CPU-only)...")
        print(f"  Command: {' '.join(cmd)}")

        try:
            result = subprocess.run(
                cmd,
                env=env,
                capture_output=True,
                text=True,
                timeout=600,  # 10 min max for large PDFs
            )
        except subprocess.TimeoutExpired:
            _warn("MinerU timed out after 10 minutes")
            return False
        except Exception as e:
            _warn(f"MinerU subprocess failed: {e}")
            return False

        if result.returncode != 0:
            _warn(f"MinerU exited with code {result.returncode}")
            if result.stderr:
                # Print first 500 chars of stderr for debugging
                _warn(f"stderr: {result.stderr[:500]}")
            return False

        # Find the output .md file in MinerU's output directory
        # MinerU creates: <tmpdir>/<pdf_stem>/auto/<pdf_stem>.md
        md_files = list(tmpdir_path.rglob("*.md"))
        if not md_files:
            _warn("MinerU produced no .md output files")
            return False

        # Take the largest .md file (most likely the full conversion)
        source_md = max(md_files, key=lambda f: f.stat().st_size)

        print(f"  MinerU output: {source_md} "
              f"({source_md.stat().st_size:,} bytes)")

        # Copy to requested output location
        output_md.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source_md, output_md)

        # Also copy any images MinerU extracted
        mineru_images = list(tmpdir_path.rglob("*.png"))
        mineru_images += list(tmpdir_path.rglob("*.jpg"))
        mineru_images += list(tmpdir_path.rglob("*.jpeg"))
        if mineru_images:
            images_dest = output_md.parent / "images" / "mineru"
            images_dest.mkdir(parents=True, exist_ok=True)
            for img in mineru_images:
                shutil.copy2(img, images_dest / img.name)
            print(f"  Copied {len(mineru_images)} images to "
                  f"{images_dest}")

        return True


def main():
    parser = argparse.ArgumentParser(
        description="Convert PDF to Markdown using MinerU (CPU-only)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
MinerU venv expected at: ~/envs/mineru/
SAFETY: CPU-only mode enforced (MPS/GPU disabled).

Examples:
  python3 convert-mineru.py scanned.pdf --output scanned.md
        """
    )

    parser.add_argument("input_file", type=Path,
                        help="Input PDF file")
    parser.add_argument("--output", "-o", type=Path, default=None,
                        help="Output markdown path "
                             "(default: <input>.md)")

    args = parser.parse_args()

    if not args.input_file.exists():
        _fail(f"Input file not found: {args.input_file}")

    if args.input_file.suffix.lower() != ".pdf":
        _fail(f"MinerU only handles PDF files, got: "
              f"{args.input_file.suffix}")

    output_md = args.output or args.input_file.with_suffix(".md")

    print(f"MinerU Converter (CPU-only)")
    print(f"Input:  {args.input_file}")
    print(f"Output: {output_md}")

    # Check MinerU is installed
    if not check_mineru_installed():
        print("\nMinerU is not installed. To install:")
        print(f"  python3 -m venv {MINERU_VENV}")
        print(f"  {MINERU_PYTHON} -m pip install magic-pdf[full]")
        sys.exit(1)

    # Run conversion
    success = convert_with_mineru(args.input_file, output_md)

    if success:
        print(f"\nConversion complete: {output_md}")
        print(f"  Size: {output_md.stat().st_size:,} bytes")
        sys.exit(0)
    else:
        _fail("MinerU conversion failed")


if __name__ == "__main__":
    main()
