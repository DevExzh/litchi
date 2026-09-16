#!/bin/bash
# Native perf stat, whole child, A1 B1 B2 A2 on CPU 21 (change 0646).
set -u
BEFORE=$1; AFTER=$2; OUT=$3; SAMPLES=$4; WARMUP=$5
CASES=pptx_cross_copy_plain,pptx_cross_copy_media_rich,pptx_cross_copy_plain_lifecycle,pptx_cross_copy_media_rich_lifecycle
run() {
  local bin=$1 tag=$2
  taskset -c 21 perf stat -e cycles,instructions -x, -o "$OUT/perf-$tag.txt" -- \
    "$bin" --warmup "$WARMUP" --samples "$SAMPLES" --case "$CASES" \
    --json "$OUT/perf-$tag.json" > "$OUT/perf-$tag.log" 2>&1
  echo "$tag exit=$?"
}
run "$BEFORE" A1
run "$AFTER"  B1
run "$AFTER"  B2
run "$BEFORE" A2
