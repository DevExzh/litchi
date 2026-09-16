#!/bin/bash
# usage: cg.sh <bin> <leg> <case> <samples> <outdir>   (change 0656, CPU 11)
set -u
BIN=$1; LEG=$2; CASE=$3; SAMPLES=$4; OUT=$5
mkdir -p "$OUT"
OUTFILE="$OUT/cg-$LEG-$CASE-s$SAMPLES.out"
taskset -c 11 valgrind --tool=callgrind --callgrind-out-file="$OUTFILE" \
  --cache-sim=no --branch-sim=no \
  "$BIN" --warmup 0 --samples "$SAMPLES" --case "$CASE" \
  --json "$OUT/cg-$LEG-$CASE-s$SAMPLES.json" \
  > "$OUT/cg-$LEG-$CASE-s$SAMPLES.runlog.txt" 2>&1
echo "callgrind exit=$?"
callgrind_annotate --threshold=99.9 "$OUTFILE" > "$OUT/incl-$LEG-$CASE-s$SAMPLES.txt" 2>&1
echo "annotate exit=$?"
