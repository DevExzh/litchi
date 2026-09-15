#!/bin/bash
# Paired timing: A1 B1 B2 A2 on CPU 21.
set -u
BEFORE=$1; AFTER=$2; OUT=$3; SAMPLES=$4; WARMUP=$5
CASES=pptx_cross_copy_plain,pptx_cross_copy_media_rich,pptx_cross_copy_plain_lifecycle,pptx_cross_copy_media_rich_lifecycle
run() {
  local bin=$1 tag=$2
  taskset -c 21 "$bin" --warmup "$WARMUP" --samples "$SAMPLES" --case "$CASES" \
    --json "$OUT/timing-$tag.json" > "$OUT/timing-$tag.log" 2>&1
  echo "$tag exit=$?"
}
run "$BEFORE" A1
run "$AFTER"  B1
run "$AFTER"  B2
run "$BEFORE" A2
