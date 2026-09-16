#!/bin/bash
# Paired timing, order A1 B1 B2 A2 in one window, pinned to CPU 15.
# usage: timing.sh <outdir> <stage-dir> <fsroot>
set -eu
OUT="$1"; STAGE="$2"; FSROOT="$3"
mkdir -p "$OUT" "$FSROOT"
CASES=docx_semantic_one_edit_save,docx_semantic_one_percent_edit_save,docx_semantic_noop_edit_save
ORD=docx_ordinary_save_lifecycle,docx_ordinary_save_edit,docx_ordinary_save_atomic_publish,docx_ordinary_save_counting_publish
run() { # run <binary> <label>
  taskset -c 15 "$STAGE/$1" --warmup 5 --samples 50 --case "$CASES" \
    --json "$OUT/$2.json" > "$OUT/$2.txt" 2>&1
  taskset -c 15 "$STAGE/$1" --warmup 5 --samples 50 --case "$ORD" \
    --filesystem-root "$FSROOT" --json "$OUT/ord-$2.json" > "$OUT/ord-$2.txt" 2>&1
}
run before A1
run after  B1
run after  B2
run before A2
echo "timing done"
