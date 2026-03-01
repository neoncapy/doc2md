#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
QC Pass 1 - Structural Check (automated).

MODIFIED from ~/.claude/scripts/qc-structural.py with 6 NEW capabilities:
1. Manifest consistency check (manifest vs image index table)
2. Image file existence check
3. Markdown syntax validation
4. Context summary JSON validation
5. Exit code semantics (0=PASS, 1=FAIL, 2=WARN)
6. Table-collapse detection (R17): detects multi-column tables collapsed
   into fewer cells during PDF conversion; inserts HTML WARNING comments
   into the .md file and reports WARN issues.

PRESERVED capabilities:
- YAML header validation
- Section/heading count
- Table column consistency
- Reference numbering [1]-[N]
- Encoding error detection
- Image index completeness

Validates the structural integrity of a converted markdown file.
This is the first gate between Python and Claude steps.

Usage:
    python3 qc-structural.py <converted.md> [--pages N] [--verbose]

Exit codes:
    0 = PASS (no issues)
    1 = FAIL (must fix before proceeding)
    2 = WARN (fix and rerun — do NOT proceed on WARN)
"""

import argparse
import json
import re
import sys
from collections import Counter
from pathlib import Path
from typing import Optional


def check_header_block(content: str) -> list[str]:
    """Verify the YAML header block has all required fields."""
    issues = []
    required_fields = [
        "source_file",
        "conversion_date",
        "conversion_tool",
        "fidelity_standard",
        "document_type",
    ]

    # Find header block
    header_match = re.match(r'^---\n(.*?)\n---', content, re.DOTALL)
    if not header_match:
        issues.append("FAIL: No YAML header block found")
        return issues

    header_text = header_match.group(1)
    for field in required_fields:
        if field + ":" not in header_text:
            issues.append(f"FAIL: Header missing required field '{field}'")

    # Conditional validation: PDF should have 'pages:', PPTX should have 'slides:'
    source_format_match = re.search(r'source_format:\s*(\w+)', header_text)
    if source_format_match:
        fmt = source_format_match.group(1).lower()
        if fmt == "pdf" and "pages:" not in header_text:
            issues.append("FAIL: PDF format but header missing 'pages:' field")
        elif fmt == "pptx" and "slides:" not in header_text:
            issues.append("FAIL: PPTX format but header missing 'slides:' field")
        # XLSX has 'sheets:', DOCX has no page count (both acceptable)

    return issues


def check_sections(content: str, expected_pages: Optional[int]) -> list[str]:
    """Check section/header structure."""
    issues = []

    # Count headings by level
    h1 = re.findall(r'^# .+', content, re.MULTILINE)
    h2 = re.findall(r'^## .+', content, re.MULTILINE)
    h3 = re.findall(r'^### .+', content, re.MULTILINE)
    total_headings = len(h1) + len(h2) + len(h3)

    if total_headings == 0:
        issues.append("FAIL: No headings found in document")
    elif total_headings < 3:
        issues.append(f"WARN: Only {total_headings} headings found (expected more for a research paper)")

    # Heuristic: expect roughly 1-3 headings per page
    if expected_pages and total_headings < expected_pages // 3:
        issues.append(
            f"WARN: Low heading count ({total_headings}) for "
            f"{expected_pages}-page document. Possible missing sections."
        )

    return issues


def check_tables(content: str) -> list[str]:
    """Verify table structure and column consistency.

    Detects BOTH pipe tables (``| col | col |``) and grid tables
    (``+-----+-----+`` borders used by pandoc's ``-t markdown`` format).
    """
    issues = []

    # --- Pipe tables (GFM style) ---
    pipe_table_blocks = []
    current_table = []
    in_table = False

    for line in content.split('\n'):
        stripped = line.strip()
        if stripped.startswith('|') and stripped.endswith('|'):
            in_table = True
            current_table.append(stripped)
        else:
            if in_table and current_table:
                pipe_table_blocks.append(current_table)
                current_table = []
            in_table = False

    if in_table and current_table:
        pipe_table_blocks.append(current_table)

    for idx, table in enumerate(pipe_table_blocks):
        col_counts = []
        for row in table:
            # Skip separator rows (|---|---|)
            if re.match(r'^\|[\s\-:]+\|$', row.strip()):
                continue
            cols = len(row.split('|')) - 2  # Subtract leading/trailing empty
            col_counts.append(cols)

        if col_counts and len(set(col_counts)) > 1:
            issues.append(
                f"FAIL: Pipe table {idx + 1} has inconsistent column counts: "
                f"{sorted(set(col_counts))}. Rows: {len(table)}"
            )

    # --- Grid tables (pandoc markdown style: +---+---+ borders) ---
    grid_table_blocks = []
    current_grid = []
    in_grid = False
    grid_border_re = re.compile(r'^\+[-=+]+\+$')

    for line in content.split('\n'):
        stripped = line.strip()
        is_border = bool(grid_border_re.match(stripped))
        # Grid tables use both +---+ borders AND | cell | rows
        if is_border or (in_grid and stripped.startswith('|')):
            if not in_grid:
                in_grid = True
            current_grid.append(stripped)
        else:
            if in_grid and current_grid:
                # Verify it's a real grid table (must have at least 2 border rows)
                border_count = sum(1 for r in current_grid if grid_border_re.match(r))
                if border_count >= 2:
                    grid_table_blocks.append(current_grid)
                current_grid = []
            in_grid = False

    if in_grid and current_grid:
        border_count = sum(1 for r in current_grid if grid_border_re.match(r))
        if border_count >= 2:
            grid_table_blocks.append(current_grid)

    for idx, table in enumerate(grid_table_blocks):
        # Check column consistency via border rows
        border_col_counts = []
        for row in table:
            if grid_border_re.match(row):
                # Count columns by counting segments between +
                cols = row.count('+') - 1
                border_col_counts.append(cols)
        if border_col_counts and len(set(border_col_counts)) > 1:
            issues.append(
                f"FAIL: Grid table {idx + 1} has inconsistent column counts: "
                f"{sorted(set(border_col_counts))}. Rows: {len(table)}"
            )

    total_tables = len(pipe_table_blocks) + len(grid_table_blocks)
    if total_tables == 0:
        issues.append("INFO: No markdown tables found")
    else:
        parts = []
        if pipe_table_blocks:
            parts.append(f"{len(pipe_table_blocks)} pipe")
        if grid_table_blocks:
            parts.append(f"{len(grid_table_blocks)} grid")
        issues.append(f"INFO: Found {total_tables} table(s) ({', '.join(parts)})")

    return issues


def _parse_table_blocks_with_positions(lines: list[str]) -> list[dict]:
    """
    Parse markdown table blocks from a list of lines.

    Returns a list of dicts, each with:
      - 'rows': list of stripped row strings
      - 'start': index of first line of this table in `lines`
      - 'end': index of last line of this table in `lines` (inclusive)
    """
    table_blocks = []
    current_rows = []
    start_idx = None

    for i, line in enumerate(lines):
        stripped = line.strip()
        if stripped.startswith('|') and stripped.endswith('|'):
            if not current_rows:
                start_idx = i
            current_rows.append(stripped)
        else:
            if current_rows:
                table_blocks.append({
                    'rows': current_rows,
                    'start': start_idx,
                    'end': i - 1,
                })
                current_rows = []
                start_idx = None

    if current_rows:
        table_blocks.append({
            'rows': current_rows,
            'start': start_idx,
            'end': len(lines) - 1,
        })

    return table_blocks


def _is_separator_row(row: str) -> bool:
    """
    Return True if `row` is a markdown table separator row (e.g. |---|---|).

    Handles multi-cell separators: each cell between pipes must contain
    only dashes, colons, and whitespace (the GFM table alignment syntax).
    """
    stripped = row.strip()
    if not (stripped.startswith('|') and stripped.endswith('|')):
        return False
    # Split by pipe, skip the empty strings from leading/trailing pipes
    cells = stripped.split('|')[1:-1]
    if not cells:
        return False
    # Every cell must contain at least one dash (not just spaces/colons)
    # and be composed only of dashes, colons, and spaces.
    # Without the dash requirement, whitespace-only cells (|   |   |)
    # would be misclassified as separators (I7 fix).
    return all(re.match(r'^[\s:]*-[\s\-:]*$', cell) for cell in cells)


def _count_numeric_values_per_cell(row: str) -> list[int]:
    """
    Count numeric tokens in each cell of a table row individually.

    Returns a list of integers, one per cell, where each integer is the
    count of numeric tokens found in that cell.

    Before counting, known multi-token patterns that represent a SINGLE
    value are neutralised:
      - ISO dates: 2026-01-15 → replaced so they don't parse as 3 tokens (I2 fix)
      - Ranges: 10-20, 2020-2025, 0.61-0.82 → replaced as single token (I3 fix)

    Numeric tokens are:
      - Integers: 0, 42, 1986
      - Decimals: 0.75, -0.07, 3.14
      - Percentages: 15%, 0.5%
      - Scientific notation: 1.2e-5
    """
    stripped = row.strip()
    if not (stripped.startswith('|') and stripped.endswith('|')):
        return []
    cells = stripped.split('|')[1:-1]  # drop leading/trailing empty strings

    # Patterns for multi-token single values (applied per cell)
    # ISO date: 2026-01-15, 2026/01/15
    date_pattern = re.compile(r'\d{4}[-/]\d{1,2}[-/]\d{1,2}')
    # Range: 10-20, 10–20, 0.61-0.82, £10,000-£15,000
    # Must have digit immediately before the dash/en-dash and digit after
    range_pattern = re.compile(r'\d[-–]\d')

    numeric_pattern = re.compile(
        r'-?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?%?'
    )

    counts = []
    for cell in cells:
        cleaned = cell
        # Strip markdown bold/italic markers before numeric counting (C2 fix).
        # Bold: **text** or *text* in academic tables (e.g. **95% CI: 0.61-0.82**)
        # cause false-positive multi-value counts when the ** surround numbers.
        # Simple replacement is safe here because we only care about numeric tokens.
        cleaned = re.sub(r'\*{1,2}', '', cleaned)
        # Fix 3.10B: Replace WxH dimension patterns (e.g., "1920x1080",
        # "640x480") with a single placeholder. These appear in Image
        # Index tables and should count as 1 value, not 2 (m8 fix).
        cleaned = re.sub(r'\d+\s*[xX]\s*\d+', '__DIM__', cleaned)
        # Replace ISO dates with a placeholder (1 value)
        cleaned = date_pattern.sub('__DATE__', cleaned)
        # For ranges: replace the dash between digits with a space
        # so the two sides don't merge, but we'll count the whole
        # range as a single token by replacing with placeholder
        # Strategy: if a range is detected, replace the entire
        # range match with a single placeholder
        # Full range pattern: number-dash-number (optionally with decimals)
        full_range = re.compile(
            r'-?\d+(?:[,.]?\d+)*%?[-–]-?\d+(?:[,.]?\d+)*%?'
        )
        cleaned = full_range.sub('__RANGE__', cleaned)

        # Count remaining numeric tokens
        nc = len(numeric_pattern.findall(cleaned))
        # Add back 1 for each date placeholder (dates are 1 value)
        nc += cleaned.count('__DATE__')
        # Add back 1 for each range placeholder (ranges are 1 value)
        nc += cleaned.count('__RANGE__')
        # Add back 1 for each dimension placeholder (dimensions are 1 value)
        nc += cleaned.count('__DIM__')
        counts.append(nc)

    return counts


def check_table_collapse(content: str) -> tuple[list[str], list[int]]:
    """
    Detect markdown tables that may have had columns collapsed during PDF
    conversion (R17).

    Detection logic:
      - For each data row in each table, count declared columns
        (pipe characters - 2, matching the behaviour in check_tables).
      - Count cells with MORE THAN 2 numeric values (m-4 fix: threshold
        raised from >1 to >2 to avoid false positives on valid tables
        where cells contain a value + parenthetical annotation, such as
        "0.85 (0.12)" for mean+SE — those have exactly 2 tokens, not 3+).
      - A row is "collapsed" if at least one of its cells has >2 numeric
        values (3+ independent numbers packed together signals data
        merged from multiple original cells).
      - If the MAJORITY (>50%) of data rows are "collapsed", flag the table.

    Conservative thresholds (to minimise false positives, m-4 fix):
      - Per-cell threshold: >2 numeric values (not >1). Cells with
        "value (annotation)" have exactly 2 tokens; cells with collapsed
        data have 3+ independent values crammed together.
      - Single-column tables with long numeric expressions (e.g. formulas
        written as prose) are excluded (1-column tables are not flagged).
      - Tables with fewer than 2 data rows are excluded (too little
        evidence).

    Returns:
      - issues: list[str] — WARN messages for each flagged table
      - flagged_end_lines: list[int] — line indices (0-based) immediately
        after the last line of each flagged table, suitable for inserting
        HTML WARNING comments.
    """
    issues = []
    flagged_end_lines = []

    lines = content.split('\n')
    table_blocks = _parse_table_blocks_with_positions(lines)

    if not table_blocks:
        return issues, flagged_end_lines

    for idx, block in enumerate(table_blocks):
        rows = block['rows']
        table_num = idx + 1

        # ── Fix 1.2: Skip tables under "## Image Index" heading ──────────
        # Image Index tables contain filenames like image-001.png whose
        # numeric tokens trigger false-positive numeric density detection.
        # Scan backward from the table start to find the nearest heading.
        start_line = block['start']
        is_image_index_table = False
        for scan_idx in range(start_line - 1, -1, -1):
            scan_line = lines[scan_idx].strip()
            if scan_line.startswith('#'):
                if 'image index' in scan_line.lower():
                    is_image_index_table = True
                break  # stop at the first heading found
        # Secondary check: if any cell contains a file extension pattern,
        # it is likely a file listing table, not a data table.
        if not is_image_index_table:
            file_ext_pattern = re.compile(
                r'\.(png|jpg|jpeg|gif|tiff|svg|bmp|webp)',
                re.IGNORECASE
            )
            ext_row_count = sum(
                1 for r in rows if file_ext_pattern.search(r)
            )
            # If majority of rows contain file extensions, skip
            if len(rows) > 0 and ext_row_count / len(rows) > 0.5:
                is_image_index_table = True
        # Fix 3.10B: Header keyword detection (m8) -- catches Image Index
        # tables even when not under an "Image Index" heading (e.g.,
        # non-standard heading text or tables embedded in other sections).
        # Requires 2+ keyword matches to avoid false skips on data tables
        # that happen to have one matching column name.
        if not is_image_index_table:
            _data_rows_for_hdr = [
                r for r in rows if not _is_separator_row(r)
            ]
            _header = _data_rows_for_hdr[0] if _data_rows_for_hdr else ""
            _index_keywords = [
                "page", "file", "classification", "dimensions",
                "figure_num", "filename", "image_path", "fig_",
                "page_num", "img_",
            ]
            _kw_hits = sum(
                1 for kw in _index_keywords
                if kw in _header.lower()
            )
            if _kw_hits >= 2:
                is_image_index_table = True
        if is_image_index_table:
            continue
        # ──────────────────────────────────────────────────────────────────

        # Collect data rows only (skip separator rows)
        data_rows = [r for r in rows if not _is_separator_row(r)]

        # Need at least 2 data rows (header + at least 1 data row) to judge
        if len(data_rows) < 2:
            continue

        # Determine declared column count from the majority of data rows
        col_counts_per_row = []
        for row in data_rows:
            declared_cols = len(row.split('|')) - 2
            col_counts_per_row.append(max(declared_cols, 1))

        col_count_majority = Counter(col_counts_per_row).most_common(1)[0][0]

        # ── Heuristic A: Zero-numeric detection (Fix A, S13) ─────────────────
        # Run BEFORE single-column guard so it catches both 1-col and
        # multi-col tables that have lost all numeric data columns.
        # Condition: (a) <= 4 total data rows, (b) zero numeric values total,
        # (c) at least 1 non-empty data row (avoids flagging empty tables).
        # Catches tables like Table 2 (col collapsed, only text labels remain)
        # and Table 11 (all data columns dropped, only label column survives).
        #
        # S15 fix: Added "well-populated text table" exemption.
        # Multi-column text-only tables (definitions, comparisons, methods)
        # are common in study guides and reference documents.  A zero-numeric
        # table is only suspicious if cells are mostly EMPTY (signals data
        # columns were dropped).  If cells are well-populated with text,
        # the table is legitimate — not collapsed.
        data_value_rows_early = data_rows[1:] if len(data_rows) > 1 else []
        if data_value_rows_early:
            total_numeric_early = sum(
                sum(_count_numeric_values_per_cell(r))
                for r in data_value_rows_early
            )
            non_empty_early = [
                r for r in data_value_rows_early
                if r.strip().replace('|', '').replace(' ', '').replace('*', '')
            ]
            if total_numeric_early == 0 and len(data_rows) <= 4 and non_empty_early:
                # S15 guard: check if this is a well-populated text table.
                # For multi-column tables (2+ cols), count cells that contain
                # meaningful text (>3 non-whitespace chars after stripping
                # pipes and markdown).  If >=60% of cells are populated,
                # this is a legitimate text table, not a collapsed data table.
                is_text_table = False
                if col_count_majority >= 2:
                    total_cells = 0
                    filled_cells = 0
                    for row in data_value_rows_early:
                        cells = row.strip().split('|')[1:-1]
                        for cell in cells:
                            total_cells += 1
                            # Strip markdown formatting before checking
                            cell_text = re.sub(r'\*{1,2}', '', cell).strip()
                            if len(cell_text) > 3:
                                filled_cells += 1
                    if total_cells > 0 and filled_cells / total_cells >= 0.6:
                        is_text_table = True

                if not is_text_table:
                    issues.append(
                        f"WARN: Table {table_num} may have collapsed columns. "
                        f"Table has {len(data_rows)} row(s) with zero numeric "
                        f"values across all data cells — expected data appears "
                        f"missing. Manual verification recommended."
                    )
                    flagged_end_lines.append(block['end'])
                    continue  # skip remaining checks; already flagged
        # ─────────────────────────────────────────────────────────────────────

        # Single-column tables: mostly skip, but flag 3+ row substantive ones
        # (Fix B, S13: catches tables where all columns except the label column
        # were dropped during DOCX-to-markdown conversion, e.g. Table 11).
        # Valid single-column use cases (note boxes, 1-2 row callouts) have
        # <= 2 substantive rows; 3+ rows is a reliable collapse signal.
        if col_count_majority <= 1:
            data_value_rows_sc = data_rows[1:] if len(data_rows) > 1 else []
            substantive_rows = [
                r for r in data_value_rows_sc
                if r.strip().replace('|', '').replace(' ', '').replace('*', '')
            ]
            if len(substantive_rows) >= 3:
                issues.append(
                    f"WARN: Table {table_num} may have collapsed columns. "
                    f"Single-column table with {len(substantive_rows)} data "
                    f"row(s) detected. Multi-column tables frequently collapse "
                    f"to 1 column during DOCX-to-markdown conversion. "
                    f"Manual verification recommended."
                )
                flagged_end_lines.append(block['end'])
            continue  # always skip single-column for numeric density check

        # ── Per-cell numeric density detection (I1 core fix, updated m-4) ──
        # A healthy table cell has 0-2 numeric values. Valid statistical
        # cells commonly hold "value (annotation)" patterns, e.g., "0.85
        # (0.12)" for mean + SE — 2 tokens, NOT a collapse signal.
        # A collapsed cell has 3+ independent values crammed together
        # (e.g., "0.10 0.05 0.07" = 3 rows merged into 1 cell).
        # Threshold is >2 (m-4 fix, raised from >1 to reduce false positives).
        #
        # Exclude header row from collapse counting entirely (I6 fix):
        # iterate only over data_value_rows (data_rows[1:]).
        data_value_rows = data_rows[1:]  # exclude header row
        if not data_value_rows:
            continue

        collapsed_row_count = 0
        max_multi_value_cells = 0

        for row in data_value_rows:
            cell_counts = _count_numeric_values_per_cell(row)
            # Count cells in this row that have >2 numeric values (m-4 fix).
            #
            # Threshold raised from >1 to >2.  Rationale:
            #
            # Valid statistical/financial tables routinely contain cells
            # with a point estimate + one parenthetical annotation, e.g.:
            #   "0.85 (0.12)"   — mean and standard error
            #   "1.5 (0.3)"     — coefficient and SE
            #   "1234 (5%)"     — revenue and growth rate
            # These cells contain exactly 2 numeric tokens and are NORMAL.
            # The old >1 threshold flagged ALL such cells as "multi-value",
            # causing false positives on every valid stats/finance table.
            #
            # A genuinely collapsed cell contains 3+ independent values
            # crammed together (e.g., multiple rows of data merged):
            #   "0.85 0.12 0.95"       — 3 independent values
            #   "1.5 0.3 2.1 0.4"      — 4 independent values
            # These clearly indicate pandoc table collapse and are what
            # we want to detect.
            multi_value_cells = sum(1 for c in cell_counts if c > 2)
            max_multi_value_cells = max(max_multi_value_cells, multi_value_cells)
            # A row is "collapsed" if at least one of its cells has >2
            # numeric values (signals data from multiple source cells
            # merged into a single markdown cell).
            if multi_value_cells > 0:
                collapsed_row_count += 1

        collapsed_fraction = collapsed_row_count / len(data_value_rows)

        # Require majority (>50%) of data rows to have at least one cell
        # with >2 numeric values to flag the table (m-4 fix).
        if collapsed_fraction > 0.5:
            # Calculate max total numeric values in any row for message
            max_values_seen = max(
                (sum(_count_numeric_values_per_cell(r))
                 for r in data_value_rows),
                default=0
            )
            issues.append(
                f"WARN: Table {table_num} may have collapsed columns. "
                f"Detected cells with multiple numeric values "
                f"({max_multi_value_cells} multi-value cell(s) in worst row) "
                f"across {col_count_majority} column(s). "
                f"{collapsed_row_count}/{len(data_value_rows)} data rows "
                f"affected. Manual verification recommended."
            )
            flagged_end_lines.append(block['end'])

    return issues, flagged_end_lines


def annotate_collapsed_tables(content: str, flagged_end_lines: list[int]) -> str:
    """
    Insert HTML WARNING comments immediately after each flagged table.

    The comment is inserted AFTER the last line of the table block.
    Uses the exact warning text specified in R17:
        <!-- WARNING: Table may have collapsed columns. Original had
        N values but only M columns detected. Manual verification
        recommended. -->

    Note: flagged_end_lines are 0-based line indices. Comments are
    inserted after those lines. Multiple insertions are handled by
    processing in reverse order to preserve line indices.

    Returns the modified content string.
    """
    if not flagged_end_lines:
        return content

    lines = content.split('\n')
    table_blocks = _parse_table_blocks_with_positions(lines)

    # Build a map from end-line index to table block for accurate messaging
    end_to_block = {b['end']: b for b in table_blocks}

    # Process in reverse order so earlier insertions don't shift indices
    for end_idx in sorted(set(flagged_end_lines), reverse=True):
        block = end_to_block.get(end_idx)
        if block is None:
            continue

        # Calculate max values and column count for the warning message
        rows = block['rows']
        data_rows = [r for r in rows if not _is_separator_row(r)]
        # Exclude header row for collapse counting (matches check_table_collapse)
        data_value_rows = data_rows[1:]
        # Use ALL data_rows for col_count_majority to match check_table_collapse
        col_counts_all = [max(len(r.split('|')) - 2, 1) for r in data_rows]
        col_count_majority = (
            Counter(col_counts_all).most_common(1)[0][0] if col_counts_all else 1
        )
        # Sum per-cell counts to get total numeric values in each row
        max_values = max(
            (sum(_count_numeric_values_per_cell(r)) for r in data_value_rows),
            default=0
        )

        comment = (
            f"<!-- WARNING: Table may have collapsed columns. "
            f"Original had {max_values} values but only "
            f"{col_count_majority} columns detected. "
            f"Manual verification recommended. -->"
        )
        # Insert the comment after the last line of the table,
        # but only if the comment is not already present nearby
        # (idempotency guard — I5 fix: scan past blank lines).
        # Fix 1.4: Extended to check BOTH forward AND backward for
        # existing warnings. Forward catches warnings after table;
        # backward catches warnings before table (e.g., manually
        # added or from prior QC runs with different table parsing).
        insert_pos = end_idx + 1
        already_annotated = False
        # Forward scan: look ahead up to 5 lines past the table end
        for lookahead in range(5):
            check_idx = insert_pos + lookahead
            if check_idx >= len(lines):
                break
            if 'WARNING: Table may have collapsed columns' in lines[check_idx]:
                already_annotated = True
                break
            # Stop scanning if we hit non-blank, non-comment content
            if lines[check_idx].strip() and not lines[check_idx].strip().startswith('<!--'):
                break
        # Backward scan: look behind up to 5 lines before the table
        # start for an existing WARNING comment (catches warnings
        # placed above the table in prior QC runs).
        if not already_annotated:
            table_start = block['start']
            for lookback in range(1, 6):
                check_idx = table_start - lookback
                if check_idx < 0:
                    break
                if 'WARNING: Table may have collapsed columns' in lines[check_idx]:
                    already_annotated = True
                    break
                # Stop scanning if we hit non-blank, non-comment content
                if lines[check_idx].strip() and not lines[check_idx].strip().startswith('<!--'):
                    break
        if not already_annotated:
            lines.insert(insert_pos, comment)

    return '\n'.join(lines)


def check_references(content: str) -> list[str]:
    """Check that reference numbers [1]-[N] are sequential."""
    issues = []

    # ── Fix 1.1: Use \d{1,4} to exclude DOI fragments (e.g. [0962280211419645])
    # that caused MemoryError when range() tried to allocate trillions of elements.
    # Academic papers never have >9999 references, so 1-4 digits is sufficient.

    # Find all bracketed single numbers like [1], [2], etc. (1-4 digits only)
    refs = re.findall(r'\[(\d{1,4})\]', content)

    # ── Fix 1.3: Parse range citations like [6-8] or [6–8] and group
    # citations like [6,9,10] that the single-number regex misses.

    # Range refs: [6-8] or [6–8] (en-dash variant)
    range_refs = re.findall(r'\[(\d{1,4})\s*[-–]\s*(\d{1,4})\]', content)
    for start_s, end_s in range_refs:
        start_n, end_n = int(start_s), int(end_s)
        # Sanity cap: max 50 elements per range to prevent edge cases
        if 0 < end_n - start_n <= 50:
            refs.extend(str(n) for n in range(start_n, end_n + 1))

    # Group refs: [6,9,10] or [6, 9, 10]
    # Fix 3.10A: Strip DOI patterns before group ref extraction to prevent
    # DOI fragments like [0962280211419645] from being split on commas
    # and parsed as reference numbers (M20 prevention).
    # Broader pattern catches both bare DOIs (10.xxxx/...) and prefixed
    # DOIs (doi: 10.xxxx/..., DOI: 10.xxxx/...).
    _content_no_doi = re.sub(
        r'(?:doi|DOI)[:\s]*10\.\d{4,}/\S+', '', content
    )
    _content_no_doi = re.sub(r'10\.\d{4,}/[^\s\]]+', '', _content_no_doi)
    group_refs = re.findall(r'\[([\d,\s]+)\]', _content_no_doi)
    for group in group_refs:
        nums = re.findall(r'\d+', group)
        if len(nums) > 1:  # only if actually a group (2+ numbers)
            # Filter to 1-4 digit numbers only (consistent with Fix 1.1)
            group_nums = [int(n) for n in nums if len(n) <= 4]
            # Fix 1.3a: Guard against false positives from comma-formatted
            # numbers (e.g. [1,000] → 1, 0) and year lists ([2020, 2021]).
            # Guard 1: all refs must be >= 1 (no ref numbered 0)
            # Guard 2: all refs must be <= 500 (no paper has 500+ refs)
            # Guard 3: group must have >= 2 valid numbers
            if len(group_nums) < 2:
                continue
            if any(n < 1 or n > 500 for n in group_nums):
                continue
            refs.extend(str(n) for n in group_nums)

    if not refs:
        issues.append("INFO: No bracketed references [N] found")
        return issues

    ref_nums = sorted(set(int(r) for r in refs))

    # ── Fix Issue C: Filter out year-like numbers (1900-2099) that appear
    # as bracketed years in academic text, e.g. [2017], [2024].
    # No academic paper has 1900+ numbered references, so this is safe.
    # Must happen BEFORE gap analysis to prevent ~2000 false gap warnings.
    ref_nums = [n for n in ref_nums if not (1900 <= n <= 2099)]

    # ── Fix 1.1 (secondary guard): Filter out any numbers exceeding 9999.
    # Belt-and-suspenders: the \d{1,4} regex should prevent this, but
    # group parsing could theoretically let through longer numbers.
    MAX_REF_NUMBER = 9999
    ref_nums = [n for n in ref_nums if n <= MAX_REF_NUMBER]
    if not ref_nums:
        issues.append(
            "INFO: No valid reference numbers found "
            "(all exceeded max threshold of 9999)"
        )
        return issues

    # Check sequential from 1
    if ref_nums[0] != 1:
        issues.append(f"WARN: References start at [{ref_nums[0]}], not [1]")

    # Check for gaps
    expected = list(range(ref_nums[0], ref_nums[-1] + 1))
    missing = set(expected) - set(ref_nums)
    if missing:
        missing_sorted = sorted(missing)
        if len(missing_sorted) > 10:
            issues.append(
                f"WARN: {len(missing_sorted)} missing references. "
                f"First gaps: {missing_sorted[:10]}"
            )
        else:
            issues.append(f"WARN: Missing references: {missing_sorted}")

    issues.append(f"INFO: References span [{ref_nums[0]}]-[{ref_nums[-1]}] ({len(ref_nums)} unique)")
    return issues


def check_encoding(content: str) -> list[str]:
    """Flag lines with encoding errors or garbled characters."""
    issues = []
    garbled_patterns = [
        r'[\x00-\x08\x0b\x0c\x0e-\x1f]',  # Control characters
        r'\ufffd',  # Unicode replacement character
    ]

    problem_lines = []
    for line_num, line in enumerate(content.split('\n'), 1):
        for pattern in garbled_patterns:
            if re.search(pattern, line):
                problem_lines.append(line_num)
                break

    if problem_lines:
        if len(problem_lines) > 10:
            issues.append(
                f"WARN: {len(problem_lines)} lines with encoding issues. "
                f"First: lines {problem_lines[:10]}"
            )
        else:
            issues.append(f"WARN: Encoding issues on lines: {problem_lines}")
    else:
        issues.append("PASS: No encoding errors detected")

    return issues


def check_image_index(content: str, md_path: Path = None) -> list[str]:
    """Verify image index table exists (embedded section or companion file)."""
    issues = []

    if "## Image Index" not in content:
        # Bug 2 fix: pipeline writes index as companion file (*-image-index.md),
        # not as an embedded section.  Check for the companion file before warning.
        if md_path is not None:
            companion_candidates = list(md_path.parent.glob("*-image-index.md"))
            if companion_candidates:
                issues.append(
                    f"INFO: Image Index companion file found: "
                    f"{companion_candidates[0].name}"
                )
                return issues
        issues.append("WARN: No Image Index section found")
        return issues

    # Find image references in markdown
    img_refs = re.findall(r'\[([^\]]+)\]\(([^)]+)\)', content)
    image_files = [ref for label, ref in img_refs
                   if any(ref.endswith(ext) for ext in ['.png', '.jpg', '.jpeg', '.gif', '.bmp'])]

    if not image_files:
        issues.append("INFO: No image file references found")
    else:
        issues.append(f"INFO: {len(image_files)} image reference(s) in index")

    return issues


# ═══════════════════════════════════════════════════════════════════════════
# NEW CHECKS
# ═══════════════════════════════════════════════════════════════════════════

def check_manifest_consistency(content: str, md_path: Path) -> list[str]:
    """
    Compare image-manifest.json against the Image Index table in MD.
    Checks: count match, filename match, page number match.
    """
    issues = []

    # Find manifest path from MD header or default location.
    # find_manifest_path now checks both PDF (image-manifest.json) and office
    # (*_manifest.json) naming patterns and returns the first existing file.
    manifest_path = find_manifest_path(content, md_path)
    if not manifest_path or not manifest_path.exists():
        issues.append("WARN: manifest not found (image-manifest.json or *_manifest.json) "
                      "— expected if document has no images")
        return issues

    try:
        with open(manifest_path, 'r', encoding='utf-8') as f:
            manifest = json.load(f)
    except Exception as e:
        issues.append(f"FAIL: Cannot read manifest: {e}")
        return issues

    manifest_images = manifest.get("images", [])
    manifest_count = len(manifest_images)

    # Parse image index table from MD (embedded section or companion file)
    index_section = extract_image_index_section(content)
    if not index_section and md_path is not None:
        # Companion file check: pipeline writes *-image-index.md separately.
        # The companion file uses "## Page-by-Page Index" (per-page rows) rather
        # than an embedded "## Image Index" (per-image rows).  When present, the
        # companion file IS the authoritative image index.  Validate count via
        # the header metadata line ("Total images detected: N") instead of
        # parsing table rows.
        companion_candidates = list(md_path.parent.glob("*-image-index.md"))
        if companion_candidates:
            try:
                companion_content = companion_candidates[0].read_text(
                    encoding="utf-8")
                # Extract "Total images detected: N" from companion header
                import re as _re
                _m = _re.search(
                    r"Total images detected:\s*(\d+)", companion_content)
                if _m:
                    companion_count = int(_m.group(1))
                    if companion_count == manifest_count:
                        issues.append(
                            "PASS: Manifest consistency check passed "
                            "(companion image index, count matches)")
                    else:
                        issues.append(
                            f"FAIL: Manifest count ({manifest_count}) != "
                            f"companion image index count ({companion_count})")
                    return issues
                # Companion exists but no count line — treat as PASS (soft)
                issues.append(
                    "INFO: Companion image index found (no count header to "
                    "validate against manifest)")
                return issues
            except Exception:
                pass
    if not index_section:
        if manifest_count > 0:
            issues.append(f"FAIL: Manifest has {manifest_count} images but Image Index section missing")
        return issues

    # Count rows in index table (exclude header and separator)
    table_rows = [line for line in index_section.split('\n')
                  if line.strip().startswith('|') and '---' not in line]
    # Separator rows already filtered out, so subtract 1 for header row only
    index_count = max(0, len(table_rows) - 1)

    # Check count match
    if manifest_count != index_count:
        issues.append(
            f"FAIL: Manifest count ({manifest_count}) != Image Index count ({index_count})"
        )

    # Check filename and page match
    index_filenames = extract_filenames_from_index(index_section)
    manifest_filenames = {img["filename"] for img in manifest_images}

    missing_in_index = manifest_filenames - index_filenames
    extra_in_index = index_filenames - manifest_filenames

    if missing_in_index:
        issues.append(f"FAIL: Filenames in manifest but not in index: {sorted(missing_in_index)}")
    if extra_in_index:
        issues.append(f"FAIL: Filenames in index but not in manifest: {sorted(extra_in_index)}")

    if not missing_in_index and not extra_in_index and manifest_count == index_count:
        issues.append("PASS: Manifest consistency check passed")

    return issues


def find_manifest_path(content: str, md_path: Path) -> Optional[Path]:
    """Find the manifest JSON path from YAML header or default location.

    Bug 3 fix: office formats write <stem>_manifest.json (e.g. report_manifest.json)
    at the document root, not image-manifest.json inside an images/ sub-directory.
    Check both naming patterns so QC does not spuriously warn on office conversions.
    Priority order:
      1. images/<short-name>/image-manifest.json  (PDF pipeline default)
      2. <stem>_manifest.json in same directory     (office pipeline pattern)
      3. Any *_manifest.json in same directory      (fallback glob)
    """
    # Try to extract images_dir from Image Index links
    img_links = re.findall(r'\[([^\]]+)\]\((images/[^)]+)\)', content)
    if img_links:
        # Extract images/<short-name>/ pattern
        first_link = img_links[0][1]
        images_dir_match = re.match(r'(images/[^/]+)', first_link)
        if images_dir_match:
            images_dir = md_path.parent / images_dir_match.group(1)
            candidate = images_dir / "image-manifest.json"
            if candidate.exists():
                return candidate

    # Default PDF path: images/<stem>/image-manifest.json
    short_name = md_path.stem
    default_dir = md_path.parent / "images" / short_name
    pdf_manifest = default_dir / "image-manifest.json"
    if pdf_manifest.exists():
        return pdf_manifest

    # Bug 3 fix: office pipeline writes <stem>_manifest.json at document root
    office_manifest = md_path.parent / f"{short_name}_manifest.json"
    if office_manifest.exists():
        return office_manifest

    # Fallback glob: any *_manifest.json in the same directory
    glob_matches = list(md_path.parent.glob("*_manifest.json"))
    if glob_matches:
        return glob_matches[0]

    # Return the PDF default path (caller checks .exists())
    return pdf_manifest


def extract_image_index_section(content: str) -> Optional[str]:
    """Extract the Image Index section from MD content."""
    match = re.search(r'## Image Index\n\n(.*?)\n---', content, re.DOTALL)
    if match:
        return match.group(1)
    # Try without --- separator
    match = re.search(r'## Image Index\n\n(.*?)(?=\n##|\Z)', content, re.DOTALL)
    if match:
        return match.group(1)
    return None


def extract_filenames_from_index(index_section: str) -> set:
    """Extract all filenames from the Image Index table."""
    filenames = set()
    # Pattern: [filename](path/to/filename)
    matches = re.findall(r'\[([^\]]+\.(?:png|jpg|jpeg|gif|bmp))\]', index_section)
    filenames.update(matches)
    return filenames


def check_image_files_exist(md_path: Path) -> list[str]:
    """
    Check that all images referenced in manifest actually exist on disk.
    """
    issues = []

    manifest_path = find_manifest_path("", md_path)  # Use default logic
    if not manifest_path or not manifest_path.exists():
        issues.append("INFO: No manifest to check file existence")
        return issues

    try:
        with open(manifest_path, 'r', encoding='utf-8') as f:
            manifest = json.load(f)
    except Exception as e:
        issues.append(f"WARN: Cannot read manifest for file check: {e}")
        return issues

    # Use images_dir field from manifest if present (office pipeline stores full path there)
    manifest_images_dir = manifest.get("images_dir")
    if manifest_images_dir:
        images_dir = Path(manifest_images_dir)
    else:
        images_dir = manifest_path.parent
    missing_files = []

    for img in manifest.get("images", []):
        filename = img.get("filename")
        if not filename:
            continue
        filepath = images_dir / filename
        if not filepath.exists():
            missing_files.append(filename)

    if missing_files:
        issues.append(f"FAIL: {len(missing_files)} image file(s) missing: {missing_files[:5]}")
    else:
        issues.append("PASS: All manifest images exist on disk")

    return issues


def check_markdown_syntax(content: str) -> list[str]:
    """
    Basic markdown syntax validation without external libraries.
    Checks for common issues like unclosed bold/italic, broken links.

    Bold/italic checking is done across the FULL document rather than
    per-line, because pandoc frequently emits multiline bold spans
    (** opens on line N, closes on line N+1). Per-line checking caused
    hundreds of false positives on DOCX conversions.

    Escaped asterisks (\\*) are stripped before counting so they are
    not mistaken for formatting markers.
    """
    issues = []

    # Strip escaped asterisks before counting markers
    cleaned_content = content.replace('\\*', '')

    # Count bold markers (**) across the entire document
    total_bold = cleaned_content.count('**')
    if total_bold % 2 != 0:
        issues.append(
            f"WARN: Document has odd number of bold markers (**): "
            f"{total_bold} found (expected even). One may be unclosed."
        )

    # Count italic markers (* not part of ** or ***) across the document.
    # Remove all ** first so we only count standalone *.
    # Then strip line-leading "* " (pandoc GFM bullet list markers) to avoid
    # false positives: a document with an odd number of *-style list items
    # would otherwise trigger a spurious WARN (M5 fix).
    no_bold = cleaned_content.replace('**', '')
    no_lists = re.sub(r'^\* ', '- ', no_bold, flags=re.MULTILINE)
    total_italic = no_lists.count('*')
    if total_italic % 2 != 0:
        issues.append(
            f"WARN: Document has odd number of italic markers (*): "
            f"{total_italic} found (expected even). One may be unclosed."
        )

    # Check for broken link syntax: [text]( without closing )
    broken_links = re.findall(r'\[([^\]]+)\]\([^\)]*$', content, re.MULTILINE)
    if broken_links:
        issues.append(f"WARN: {len(broken_links)} potentially broken link(s) found")

    # Check for heading hierarchy violations (H3 without preceding H2)
    # This is a simple check; not perfect for complex documents
    lines = content.split('\n')
    prev_heading_level = 0
    for line_num, line in enumerate(lines, 1):
        heading_match = re.match(r'^(#{1,6})\s', line)
        if heading_match:
            level = len(heading_match.group(1))
            if level > prev_heading_level + 1 and prev_heading_level > 0:
                issues.append(
                    f"WARN: Line {line_num} jumps from H{prev_heading_level} to H{level} "
                    f"(skipped H{prev_heading_level + 1})"
                )
            prev_heading_level = level

    if not issues:
        issues.append("PASS: Markdown syntax validation passed")

    return issues


def check_context_summary(md_path: Path, source_format: str = "") -> list[str]:
    """
    Validate context-summary.json if it exists.
    Checks for required fields: title, sections, total_images.

    context-summary.json is only produced by convert-paper.py (PDF).
    For DOCX/PPTX sources, its absence is expected and not a warning.
    """
    issues = []

    summary_path = md_path.parent / "context-summary.json"
    if not summary_path.exists():
        # Only warn for PDF sources — DOCX/PPTX never produce this file
        if source_format.lower() in ("docx", "pptx", "doc", "ppt", "txt"):
            issues.append("INFO: context-summary.json not applicable for this format")
        else:
            issues.append("WARN: context-summary.json not found (expected from convert-paper.py v2)")
        return issues

    try:
        with open(summary_path, 'r', encoding='utf-8') as f:
            summary = json.load(f)
    except Exception as e:
        issues.append(f"WARN: Cannot read context-summary.json: {e}")
        return issues

    required_fields = ["title", "sections", "total_images", "document_domain"]
    missing_fields = [f for f in required_fields if f not in summary]

    if missing_fields:
        issues.append(f"WARN: context-summary.json missing fields: {missing_fields}")
    else:
        issues.append("PASS: context-summary.json valid")

    return issues


def _record_conversion_issues(
    md_path: Path,
    collapse_issues: list[str],
) -> None:
    """
    Append flagged table-collapse warnings to CONVERSION-ISSUES.md (I4 fix).

    Searches for CONVERSION-ISSUES.md by walking up from the markdown file's
    directory until a project root is found (indicated by CLAUDE.md) or the
    filesystem root is reached. If the file is not found, creation is
    skipped silently (the project may not use this convention).

    Only WARN-level collapse issues are recorded. Each entry includes the
    source markdown path and the warning text.
    """
    warn_issues = [i for i in collapse_issues if i.startswith('WARN')]
    if not warn_issues:
        return

    # Walk up to find CONVERSION-ISSUES.md in the project tree
    conv_issues_path = None
    search_dir = md_path.parent.resolve()
    for _ in range(20):  # safety limit
        candidate = search_dir / 'CONVERSION-ISSUES.md'
        if candidate.exists():
            conv_issues_path = candidate
            break
        # Stop at project root (has CLAUDE.md) or filesystem root
        if (search_dir / 'CLAUDE.md').exists():
            # Project root found but no CONVERSION-ISSUES.md — create it here
            conv_issues_path = candidate
            break
        parent = search_dir.parent
        if parent == search_dir:
            break  # filesystem root
        search_dir = parent

    if conv_issues_path is None:
        return

    # Fix 3.10C: Deduplication guard (m11). Before appending, read the
    # existing file and check whether each issue is already recorded.
    # Signature = issue_type prefix + first 50 chars of the issue text.
    # This prevents duplicate entries when QC is run multiple times on
    # the same file (e.g., during fix-and-rerun cycles).
    existing_content = ""
    if conv_issues_path.exists():
        try:
            existing_content = conv_issues_path.read_text(encoding='utf-8')
        except OSError:
            pass  # if we can't read, proceed without dedup

    # Filter out issues already present in the file
    new_issues = []
    for issue in warn_issues:
        sig = f"{issue[:50]}"
        if sig not in existing_content:
            new_issues.append(issue)

    if not new_issues:
        return  # all issues already recorded; skip duplicate append

    from datetime import datetime
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M')

    entry_lines = [
        f"\n## Table Collapse Warning — {md_path.name} ({timestamp})\n",
        f"Source: `{md_path}`\n",
    ]
    for issue in new_issues:
        entry_lines.append(f"- {issue}\n")

    try:
        with open(conv_issues_path, 'a', encoding='utf-8') as f:
            f.writelines(entry_lines)
    except OSError:
        pass  # non-critical; don't fail the QC run


# ═══════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(
        description="QC Pass 1: Structural validation of converted markdown"
    )
    parser.add_argument("markdown_file", type=Path, help="Path to the markdown file")
    parser.add_argument(
        "--pages", "-p", type=int, default=None,
        help="Expected page count of source (for heuristic checks)"
    )
    parser.add_argument(
        "--verbose", "-v", action="store_true",
        help="Show all checks including passing ones"
    )

    args = parser.parse_args()

    if not args.markdown_file.exists():
        print(f"ERROR: File not found: {args.markdown_file}")
        sys.exit(1)

    content = args.markdown_file.read_text(encoding="utf-8")

    # Extract page count from header if not provided
    if args.pages is None:
        pages_match = re.search(r'^(?:pages|slides|sheets):\s*(\d+)', content, re.MULTILINE)
        if pages_match:
            args.pages = int(pages_match.group(1))

    # Extract source_format from YAML header for format-specific checks
    source_format = ""
    fmt_match = re.search(r'^source_format:\s*["\']?(\w+)["\']?', content, re.MULTILINE)
    if fmt_match:
        source_format = fmt_match.group(1)

    print(f"QC Pass 1: Structural Check")
    print(f"File: {args.markdown_file}")
    print(f"Size: {len(content):,} chars, {content.count(chr(10)):,} lines")
    if args.pages:
        print(f"Expected pages: {args.pages}")
    print("=" * 55)

    all_issues = []

    # ── R17: Table-collapse detection ────────────────────────────────────────
    # Run before the main checks list so we can annotate the file first.
    # check_table_collapse has a different return signature: (issues, end_lines)
    collapse_issues, flagged_end_lines = check_table_collapse(content)
    if flagged_end_lines:
        # Annotate the .md file on disk with HTML WARNING comments.
        # This modifies `content` in memory first, then writes to disk.
        annotated_content = annotate_collapsed_tables(content, flagged_end_lines)
        try:
            if annotated_content != content:
                # Content changed: new comments were inserted
                args.markdown_file.write_text(annotated_content, encoding="utf-8")
                # Count newly inserted comments
                new_count = annotated_content.count(
                    'WARNING: Table may have collapsed columns'
                ) - content.count('WARNING: Table may have collapsed columns')
                if new_count > 0:
                    print(
                        f"\nINFO: Annotated {new_count} collapsed-table "
                        f"warning(s) into {args.markdown_file}"
                    )
                content = annotated_content  # use annotated version for remaining checks
            else:
                # Content unchanged: comments already present (idempotent re-run)
                if args.verbose:
                    print(
                        f"\nINFO: Table-collapse annotations already present "
                        f"(idempotent re-run); no changes written."
                    )
        except OSError as e:
            collapse_issues.append(
                f"WARN: Could not write table-collapse annotations to file: {e}"
            )

    # R17 acceptance criterion 3: record flagged tables in CONVERSION-ISSUES.md (I4 fix)
    if collapse_issues:
        _record_conversion_issues(args.markdown_file, collapse_issues)
    # ─────────────────────────────────────────────────────────────────────────

    checks = [
        ("Header Block", check_header_block(content)),
        ("Sections", check_sections(content, args.pages)),
        ("Tables", check_tables(content)),
        ("References", check_references(content)),
        ("Encoding", check_encoding(content)),
        ("Image Index", check_image_index(content, args.markdown_file)),
        ("Manifest Consistency", check_manifest_consistency(content, args.markdown_file)),
        ("Image Files Exist", check_image_files_exist(args.markdown_file)),
        ("Markdown Syntax", check_markdown_syntax(content)),
        ("Context Summary", check_context_summary(args.markdown_file, source_format)),
        # R17: table-collapse issues injected from pre-check above
        ("Table Collapse (R17)", collapse_issues),
    ]

    fail_count = 0
    warn_count = 0

    for check_name, issues in checks:
        print(f"\n--- {check_name} ---")
        for issue in issues:
            if issue.startswith("FAIL"):
                fail_count += 1
                print(f"  {issue}")
            elif issue.startswith("WARN"):
                warn_count += 1
                print(f"  {issue}")
            elif args.verbose or issue.startswith("PASS"):
                print(f"  {issue}")
            elif issue.startswith("INFO"):
                print(f"  {issue}")
        all_issues.extend(issues)

    print("\n" + "=" * 55)
    print(f"RESULT: {fail_count} FAIL, {warn_count} WARN")

    if fail_count > 0:
        print("STATUS: NEEDS FIXES (run QC Pass 2 after fixing)")
        sys.exit(1)  # FAIL
    elif warn_count > 0:
        print("STATUS: WARN — fix and rerun (do NOT proceed to next step)")
        # NOTE: WARN exit code (2) should trigger a fix-and-rerun cycle
        # in the pipeline operator. Pipeline should NOT proceed on WARN.
        # Only PASS (exit 0) allows the pipeline to continue.
        sys.exit(2)  # WARN
    else:
        print("STATUS: PASS (proceed to next step)")
        sys.exit(0)  # PASS


if __name__ == "__main__":
    main()
