#!/bin/bash
# Deterministic instruction/call counts for one leg, by callgrind isolation pair.
# usage: counts.sh <binary> <leg-tag> <outdir>
set -u
BIN="$1"; LEG="$2"; OUT="$3"
RAW=/home/zhuhe/code/litchi-worktrees/targets/0591-cg
mkdir -p "$OUT" "$RAW"
for case in docx_semantic_one_edit_save docx_semantic_one_percent_edit_save docx_semantic_noop_edit_save; do
  for n in 1 3; do
    o="$RAW/$LEG-$case-$n.out"
    rm -f "$o" "$OUT/$LEG-$case-$n.json"
    taskset -c 11 valgrind --tool=callgrind --callgrind-out-file="$o" \
      "$BIN" --warmup 0 --samples "$n" --case "$case" \
      --json "$OUT/$LEG-$case-$n.json" > "$OUT/cg-$LEG-$case-$n.log" 2>&1
    callgrind_annotate --threshold=99.9 "$o" > "$OUT/incl-$LEG-$case-$n.txt" 2>&1
  done
done
echo "counts done for $LEG"
