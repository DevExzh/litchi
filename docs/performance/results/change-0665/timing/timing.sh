#!/bin/bash
# Change 0665 paired timing, order A1 B1 B2 A2 A3 A4 in one window, CPU 20.
# A3/A4 is the A/A floor; B1/B2 is the B/B floor.
# usage: timing.sh <outdir> <stage-dir>
set -eu
OUT="$1"; STAGE="$2"; mkdir -p "$OUT"
CASES=docx_semantic_one_edit_save,xlsx_eager_cell_values_one_edit_save,pptx_eager_batch_edit_save
run() { # run <binary> <label>
  taskset -c 20 "$STAGE/$1" --warmup 5 --samples 40 --case "$CASES" \
    --json "$OUT/$2.json" > "$OUT/$2.txt" 2>&1
}
run litchi-perf-baseline-before A1
run litchi-perf-baseline-after  B1
run litchi-perf-baseline-after  B2
run litchi-perf-baseline-before A2
run litchi-perf-baseline-before A3
run litchi-perf-baseline-before A4
echo "timing done"
