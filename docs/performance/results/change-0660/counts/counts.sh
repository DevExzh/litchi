#!/bin/bash
# Deterministic instruction/call counts for one leg, by callgrind isolation pair.
# usage: counts.sh <binary> <leg-tag> <outdir> <rawdir>
set -u
BIN="$1"; LEG="$2"; OUT="$3"; RAW="$4"
mkdir -p "$OUT" "$RAW"
for case in docx_semantic_one_edit_save docx_semantic_one_percent_edit_save docx_semantic_noop_edit_save; do
  for n in 1 3; do
    o="$RAW/$LEG-$case-$n.out"
    rm -f "$o" "$OUT/$LEG-$case-$n.json"
    taskset -c 15 valgrind --tool=callgrind --callgrind-out-file="$o" \
      "$BIN" --warmup 0 --samples "$n" --case "$case" \
      --json "$OUT/$LEG-$case-$n.json" > "$OUT/cg-$LEG-$case-$n.txt" 2>&1
    callgrind_annotate --inclusive=yes --threshold=100 "$o" | head -200 > "$OUT/incl-$LEG-$case-$n.txt" 2>&1
  done
done
echo "counts done for $LEG"
