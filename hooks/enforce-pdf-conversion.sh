#!/bin/bash
# PreToolUse hook: intercept PDF and Office document reads, redirect to markdown.
# PDF path: queries conversion_registry.json by SHA-256 hash, falls back to co-located .md.
# DOCX/PPTX/XLSX path: queries conversion_registry.json by SHA-256 hash first,
#   then falls back to co-located .md (same logic as PDF).
# Adds router-aware messaging, timestamped logging, and graceful degradation.
#
# Responsibility: GATE only (read-time enforcement).
# Registry is POPULATED by convert-office.py (office formats) or run-pipeline.py (PDF).
# This hook and those scripts interact only through the registry JSON.

INPUT=$(cat)
FILE_PATH=$(echo "$INPUT" | jq -r '.tool_input.file_path // empty')
# NOTE: TIMESTAMP is intentionally set at startup for the top-level messages.
# Individual log_event() calls recompute their own timestamp internally so that
# each log entry reflects the actual time of the event, not hook startup time.
REGISTRY="$HOME/.claude/pipeline/conversion_registry.json"
LOG_DIR="$HOME/.claude/pipeline"
LOG_FILE="$LOG_DIR/hook-interceptions.log"

# ── Helper: append a log entry (warn-only; never blocks execution) ──────────
log_event() {
  local decision="$1"
  local outcome="$2"
  # Recompute timestamp at the moment of the log event (M7: per-event accuracy)
  local LOG_TIMESTAMP
  LOG_TIMESTAMP=$(date -u +"%Y-%m-%dT%H:%M:%SZ")
  # Ensure log dir exists; silently skip if not writable
  mkdir -p "$LOG_DIR" 2>/dev/null || return 0
  # Quote FILE_PATH to handle spaces and = characters in paths (M6)
  printf '[%s] file="%s" decision=%s outcome=%s\n' \
    "$LOG_TIMESTAMP" "$FILE_PATH" "$decision" "$outcome" \
    >> "$LOG_FILE" 2>/dev/null || true
}

# ── Helper: router-aware extractor hint (no computation; heuristic only) ────
# Returns a hint string based on available tools. Does NOT run the router.
# The router (run-pipeline.py select_extractor) runs during conversion, not here.
# NOTE M8: The python3 import check runs on every blocked call. This adds
# ~200-500ms latency on slow systems. The result is cached in a flag file
# under /tmp to avoid repeated interpreter spawns within the same session.
extractor_hint() {
  local hint="pymupdf4llm (digital)"
  # Use a global cache file to avoid repeated python3 spawns within the session
  local FITZ_GLOBAL_CACHE="/tmp/hook-fitz-available"
  if [[ -f "$FITZ_GLOBAL_CACHE" ]]; then
    local cached
    cached=$(cat "$FITZ_GLOBAL_CACHE" 2>/dev/null)
    if [[ "$cached" == "0" ]]; then
      hint="unknown (PyMuPDF not found)"
    fi
  else
    if python3 -c "import fitz" 2>/dev/null; then
      echo "1" > "$FITZ_GLOBAL_CACHE" 2>/dev/null || true
    else
      hint="unknown (PyMuPDF not found)"
      echo "0" > "$FITZ_GLOBAL_CACHE" 2>/dev/null || true
    fi
  fi
  # Check if run-pipeline.py is available
  if [[ ! -f "$HOME/.claude/scripts/run-pipeline.py" ]]; then
    hint="WARNING: run-pipeline.py not found — convert-paper.py direct mode"
  fi
  echo "$hint"
}

# ── Only act on PDF file paths ────────────────────────────────────────────────
if [[ "$FILE_PATH" == *.pdf ]] || [[ "$FILE_PATH" == *.PDF ]]; then

  # Compute startup timestamp for top-level messages (log_event uses its own)
  TIMESTAMP=$(date -u +"%Y-%m-%dT%H:%M:%SZ")

  # ── Graceful degradation: warn if jq is missing ─────────────────────────
  if ! command -v jq &>/dev/null; then
    echo "WARN [hook]: jq not found. PDF interception disabled. Install jq to enable." >&2
    log_event "degraded-no-jq" "warn-only"
    exit 0
  fi

  # ── File existence check ─────────────────────────────────────────────────
  if [[ ! -f "$FILE_PATH" ]]; then
    echo "BLOCKED [$TIMESTAMP]: PDF file not found at $FILE_PATH." >&2
    log_event "file-not-found" "blocked"
    exit 2
  fi

  # ── Compute SHA-256 for registry lookup ─────────────────────────────────
  # M1: capture hash separately; check for empty result before using it.
  # shasum can fail if file is unreadable or becomes unavailable (TOCTOU).
  # On failure: warn, skip registry lookup, fall through to co-located check.
  HASH_RAW=$(shasum -a 256 "$FILE_PATH" 2>/dev/null | cut -d' ' -f1)
  if [[ -z "$HASH_RAW" ]]; then
    echo "WARN [$TIMESTAMP]: Could not compute SHA-256 for $FILE_PATH." \
         "Skipping registry lookup, checking co-located fallback." >&2
    log_event "hash-failed" "degraded"
    # Fall through to co-located check without registry lookup
  else
    SOURCE_HASH="sha256:$HASH_RAW"

    # ── Registry lookup (query by source hash) ────────────────────────────
    if [[ -f "$REGISTRY" ]]; then
      MD_PATH=$(jq -r --arg hash "$SOURCE_HASH" \
        '.conversions[] | select(.source_hash == $hash) | .output_path // empty' \
        "$REGISTRY" 2>/dev/null | head -1)

      # M2: also verify MD_PATH has a .md extension to guard against
      # corrupted registry entries pointing to non-markdown files.
      if [[ -n "$MD_PATH" && -f "$MD_PATH" && "$MD_PATH" == *.md ]]; then
        EXTRACTOR_USED=$(jq -r --arg hash "$SOURCE_HASH" \
          '.conversions[] | select(.source_hash == $hash) | .extractor_used // "unknown"' \
          "$REGISTRY" 2>/dev/null | head -1)
        echo "INFO [$TIMESTAMP]: Converted markdown found at $MD_PATH" \
             "(extractor: $EXTRACTOR_USED, matched by SHA-256). Read that instead." >&2
        log_event "registry-hit" "allowed-$MD_PATH"
        exit 0
      fi
    fi
  fi

  # ── Fallback: co-located .md (for PDFs not yet in registry) ─────────────
  MD_COLOCATED="${FILE_PATH%.pdf}.md"
  # Handle uppercase .PDF extension
  [[ "$FILE_PATH" == *.PDF ]] && MD_COLOCATED="${FILE_PATH%.PDF}.md"
  if [[ -f "$MD_COLOCATED" ]]; then
    echo "INFO [$TIMESTAMP]: Converted markdown found at $MD_COLOCATED (co-located fallback). Read that instead." >&2
    log_event "colocated-hit" "allowed-$MD_COLOCATED"
    exit 0
  fi

  # ── No conversion found: block and provide router-aware guidance ─────────
  # M3: Block with exit 2 regardless of tool availability.
  # Tool availability only affects the user-facing message, not the exit code.
  # Reading raw PDF when no converted markdown exists violates MD-First Rule R6.
  HINT=$(extractor_hint)

  if [[ ! -f "$HOME/.claude/scripts/run-pipeline.py" ]]; then
    if [[ ! -f "$HOME/.claude/scripts/convert-paper.py" ]]; then
      # Both tools missing: critical degradation — still block
      echo "BLOCKED [$TIMESTAMP]: No converted markdown found for this PDF." >&2
      echo "  CRITICAL: No conversion tools found at ~/.claude/scripts/" >&2
      echo "  Reinstall run-pipeline.py and convert-paper.py to enable conversion." >&2
      echo "  PDF: $FILE_PATH" >&2
      log_event "blocked-no-tools" "blocked-critical"
    else
      # run-pipeline.py missing but convert-paper.py available: fallback mode
      echo "BLOCKED [$TIMESTAMP]: No converted markdown found for this PDF." >&2
      echo "  WARN: run-pipeline.py not found at ~/.claude/scripts/run-pipeline.py" >&2
      # M4: use printf %q so paths with single quotes or spaces are safe to copy-paste
      printf "  FALLBACK: Convert using: python3 ~/.claude/scripts/convert-paper.py %q\n" \
             "$FILE_PATH" >&2
      echo "  PDF: $FILE_PATH" >&2
      log_event "blocked-no-pipeline" "blocked-degraded"
    fi
    exit 2
  fi

  # Normal block: run-pipeline.py available, conversion simply has not been run yet
  echo "BLOCKED [$TIMESTAMP]: No converted markdown found for this PDF." >&2
  echo "  Suggested extractor chain: $HINT" >&2
  # M4: use printf %q so paths with single quotes or spaces are safe to copy-paste
  printf "  Convert using: python3 ~/.claude/scripts/run-pipeline.py %q\n" \
         "$FILE_PATH" >&2
  echo "  (run-pipeline.py will auto-detect digital vs scanned and select extractor)" >&2
  echo "  PDF: $FILE_PATH" >&2
  log_event "blocked-no-conversion" "blocked"
  exit 2

fi

# ── Office format interception (DOCX / PPTX / XLSX) ──────────────────────────
# Registry lookup: queries conversion_registry.json by SHA-256 hash (same as PDF).
# Logic: registry hit → allow; co-located .md → allow; no .md → block with command.
# Case-insensitive: handles .docx/.DOCX, .pptx/.PPTX, .xlsx/.XLSX.
#
# Registry is populated by convert-office.py (or run-pipeline.py) on successful
# conversion. The source_hash field uses "sha256:HEXHASH" format, matching the
# jq query pattern used for PDFs above.

OFFICE_EXT=""
MD_COLOCATED_OFFICE=""

if [[ "$FILE_PATH" == *.docx ]] || [[ "$FILE_PATH" == *.DOCX ]]; then
  OFFICE_EXT="docx"
  BASE="${FILE_PATH%.*}"
  MD_COLOCATED_OFFICE="${BASE}.md"
elif [[ "$FILE_PATH" == *.pptx ]] || [[ "$FILE_PATH" == *.PPTX ]]; then
  OFFICE_EXT="pptx"
  BASE="${FILE_PATH%.*}"
  MD_COLOCATED_OFFICE="${BASE}.md"
elif [[ "$FILE_PATH" == *.xlsx ]] || [[ "$FILE_PATH" == *.XLSX ]]; then
  OFFICE_EXT="xlsx"
  BASE="${FILE_PATH%.*}"
  MD_COLOCATED_OFFICE="${BASE}.md"
fi

if [[ -n "$OFFICE_EXT" ]]; then

  TIMESTAMP=$(date -u +"%Y-%m-%dT%H:%M:%SZ")

  # ── File existence check ─────────────────────────────────────────────────
  if [[ ! -f "$FILE_PATH" ]]; then
    echo "BLOCKED [$TIMESTAMP]: Office file not found at $FILE_PATH." >&2
    log_event "office-file-not-found" "blocked"
    exit 2
  fi

  # ── Graceful degradation: warn if jq is missing (m7 fix) ────────────────
  # When jq is absent: skip registry lookup silently and fall through to the
  # co-located .md check. If no .md is found there either, the file is blocked
  # (exit 2) — same outcome as when jq is present. We now emit an explicit
  # warning so the asymmetry with the PDF block (lines 74-78) is resolved.
  if ! command -v jq &>/dev/null; then
    echo "WARN [hook]: jq not found. Office format registry lookup disabled. " \
         "Falling back to co-located .md check." >&2
    log_event "office-degraded-no-jq" "warn-only"
    # Fall through to co-located check below (no registry lookup)
  fi

  # ── Registry lookup (SHA-256, same pattern as PDF) ───────────────────────
  # shasum is always available on macOS; failure is graceful (skip to fallback).
  OFFICE_HASH_RAW=$(shasum -a 256 "$FILE_PATH" 2>/dev/null | cut -d' ' -f1)
  if [[ -n "$OFFICE_HASH_RAW" ]] && command -v jq &>/dev/null && [[ -f "$REGISTRY" ]]; then
    OFFICE_SOURCE_HASH="sha256:$OFFICE_HASH_RAW"
    OFFICE_MD_PATH=$(jq -r --arg hash "$OFFICE_SOURCE_HASH" \
      '.conversions[] | select(.source_hash == $hash) | .output_path // empty' \
      "$REGISTRY" 2>/dev/null | head -1)

    # Validate: non-empty path, file exists on disk, has .md extension
    if [[ -n "$OFFICE_MD_PATH" && -f "$OFFICE_MD_PATH" && "$OFFICE_MD_PATH" == *.md ]]; then
      OFFICE_EXTRACTOR=$(jq -r --arg hash "$OFFICE_SOURCE_HASH" \
        '.conversions[] | select(.source_hash == $hash) | .extractor_used // "unknown"' \
        "$REGISTRY" 2>/dev/null | head -1)
      echo "INFO [$TIMESTAMP]: Converted markdown found at $OFFICE_MD_PATH" \
           "(extractor: $OFFICE_EXTRACTOR, matched by SHA-256). Read that instead." >&2
      log_event "office-registry-hit" "allowed-$OFFICE_MD_PATH"
      exit 0
    fi
  fi

  # ── Fallback: co-located .md (same dir, same basename) ───────────────────
  if [[ -f "$MD_COLOCATED_OFFICE" ]]; then
    echo "INFO [$TIMESTAMP]: Converted markdown found at $MD_COLOCATED_OFFICE (co-located fallback). Read that instead." >&2
    log_event "office-colocated-hit" "allowed-$MD_COLOCATED_OFFICE"
    exit 0
  fi

  # ── No conversion found: block and suggest conversion command ─────────────
  OFFICE_EXT_UPPER=$(echo "$OFFICE_EXT" | tr '[:lower:]' '[:upper:]')
  echo "BLOCKED [$TIMESTAMP]: No converted markdown found for this ${OFFICE_EXT_UPPER} file." >&2
  echo "  MD-First Rule: read the converted .md, not the raw Office file." >&2

  case "$OFFICE_EXT" in
    docx|pptx)
      printf "  Convert using: python3 ~/.claude/scripts/run-pipeline.py %q\n" \
             "$FILE_PATH" >&2
      ;;
    xlsx)
      echo "  XLSX files should NOT be converted to markdown (lossy)." >&2
      echo "  Use Python + openpyxl to read directly: import openpyxl; wb = openpyxl.load_workbook('$FILE_PATH')" >&2
      ;;
  esac

  echo "  After conversion, read the registered .md path." >&2
  echo "  File: $FILE_PATH" >&2
  log_event "office-blocked-no-conversion" "blocked"
  exit 2

fi

# NOTE M9: update_registry() in run-pipeline.py does not write to this log file
# by design. The registry JSON is the shared interface between run-pipeline.py
# and this hook. This log file is the hook's private output only. The QC
# checklist item "Logging to conversion_registry.json" is satisfied by
# update_registry() writing to the registry JSON — not to this log.

exit 0
