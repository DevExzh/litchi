#!/bin/bash
# Change 0665 deterministic instruction/call counts, by callgrind isolation pair.
# usage: counts.sh <binary> <leg-tag> <outdir> <rawdir>
set -u
BIN="$1"; LEG="$2"; OUT="$3"; RAW="$4"
mkdir -p "$OUT" "$RAW"
for case in docx_semantic_one_edit_save xlsx_eager_cell_values_one_edit_save pptx_eager_batch_edit_save; do
  for n in 1 3; do
    o="$RAW/$LEG-$case-$n.out"
    rm -f "$o" "$OUT/$LEG-$case-$n.json"
    taskset -c 20 valgrind --tool=callgrind --callgrind-out-file="$o" \
      "$BIN" --warmup 0 --samples "$n" --case "$case" \
      --json "$OUT/$LEG-$case-$n.json" > "$OUT/cg-$LEG-$case-$n.txt" 2>&1
    callgrind_annotate --inclusive=yes --threshold=100 "$o" | head -220 > "$OUT/incl-$LEG-$case-$n.txt" 2>&1
  done
done
echo "counts done for $LEG"
