#!/bin/bash
# Isolation-pair instruction counts for the commit region alone.
# usage: region.sh <probe-binary> <leg-tag> <outdir> <rawdir>
set -u
BIN="$1"; LEG="$2"; OUT="$3"; RAW="$4"
mkdir -p "$OUT" "$RAW"
for shape in compact noncompact; do
  for paragraphs in 24 200 10000; do
    for n in 1 3; do
      o="$RAW/region-$LEG-$shape-$paragraphs-$n.out"
      rm -f "$o"
      taskset -c 15 valgrind --tool=callgrind --callgrind-out-file="$o" \
        "$BIN" commit "$paragraphs" "$shape" "$n" \
        > "$OUT/region-$LEG-$shape-$paragraphs-$n.txt" 2>&1
    done
  done
done
echo "region done for $LEG"
