#!/usr/bin/env bash
# Change 0601: descriptive baseline for the opt-in producer-shape selectors.
#
# Deterministic facts first (generator byte identity and the marker/refusal
# census, both from separate processes), then timing: an A/A pair of identical
# legs in the same window so the run states its own floor.
#
# Usage: run-baseline.sh <worktree> <output-dir> <cpu>
set -euo pipefail

WORKTREE="${1:?worktree}"
OUT="${2:?output directory}"
CPU="${3:-23}"
BIN="$WORKTREE/tools/perf-baseline/target/release/litchi-perf-baseline"
REAL_FILE="$WORKTREE/test-data/poi/test-data/spreadsheet/Excel_file_with_trash_item.xlsx"

GENERATED_CASES=(
  xlsx_producer_medium_source_open
  xlsx_producer_medium_source_selected_cell
  xlsx_producer_medium_source_planning
  xlsx_producer_medium_source_one_edit_save
  xlsx_producer_medium_control_selected_cell
  xlsx_producer_medium_control_planning
  xlsx_producer_dense_source_open
  xlsx_producer_dense_source_selected_cell
  xlsx_producer_dense_source_planning
  xlsx_producer_dense_source_one_edit_save
  xlsx_producer_dense_control_selected_cell
  xlsx_producer_dense_control_planning
  docx_producer_source_selected_paragraph
  pptx_producer_source_selected_slide
)
GENERATED=$(IFS=,; echo "${GENERATED_CASES[*]}")
REAL='xlsx_real_file_source_open,xlsx_real_file_source_selected_cell'

mkdir -p "$OUT"

# Provenance.
{
  echo "host: $(uname -srm)"
  echo "cpu: $CPU"
  echo "date_utc: $(date -u +%Y-%m-%dT%H:%M:%SZ)"
  echo "git_commit: $(git -C "$WORKTREE" rev-parse HEAD)"
  echo "git_branch: $(git -C "$WORKTREE" rev-parse --abbrev-ref HEAD)"
  echo "rustc: $(rustc --version)"
  echo "binary_sha256: $(sha256sum "$BIN" | cut -d' ' -f1)"
  echo "binary_bytes: $(stat -c %s "$BIN")"
  echo "real_file_sha256: $(sha256sum "$REAL_FILE" | cut -d' ' -f1)"
} > "$OUT/provenance.txt"

# 1. Determinism: two independent processes must emit byte-identical corpora.
for leg in 1 2; do
  taskset -c "$CPU" "$BIN" --warmup 0 --samples 1 --case "$GENERATED" \
    --json "$OUT/determinism-$leg.json" \
    --producer-evidence "$OUT/evidence-$leg.json" > /dev/null
done
diff -u "$OUT/evidence-1.json" "$OUT/evidence-2.json" > "$OUT/determinism.diff" || true
for leg in 1 2; do
  python3 "$(dirname "$0")/corpus_identity.py" \
    "$OUT/determinism-$leg.json" "$OUT/corpus-identity-$leg.json"
done
diff -u "$OUT/corpus-identity-1.json" "$OUT/corpus-identity-2.json" \
  > "$OUT/corpus-identity.diff" || true

# 2. Baseline: 20 warm-ups, 100 samples, A/A pair in the same window.
for leg in a b; do
  taskset -c "$CPU" "$BIN" --warmup 20 --samples 100 --case "$GENERATED" \
    --json "$OUT/baseline-$leg.json" > /dev/null
done

# 3. The same read scenarios over a real Excel file.
taskset -c "$CPU" "$BIN" --warmup 20 --samples 100 --case "$REAL" \
  --real-file "$REAL_FILE" \
  --json "$OUT/real-file.json" \
  --producer-evidence "$OUT/real-file-evidence.json" > /dev/null

echo "done"
