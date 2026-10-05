#!/bin/bash
# Interleaved before/after open-route loop for record 0012.
# usage: abba.sh <cpu> <rounds> <out.jsonl>
set -euo pipefail
S=/tmp/claude-1001/-home-zhuhe-code-litchi/b8eb3d11-f223-4205-92d4-66b12eacbd85/scratchpad/0012
T=/home/zhuhe/code/litchi-worktrees/uo-0012-opc-shared-package/test-data
CPU=$1; ROUNDS=$2; OUT=$3
# route fixture warmup samples
SPECS=(
  "opc-open ooxml/docx/drawing.docx 200 2000"
  "docx-open ooxml/docx/drawing.docx 200 2000"
  "docx-open-text ooxml/docx/drawing.docx 50 400"
  "docx-open-text ooxml/docx/comment.docx 200 2000"
  "opc-open poi/test-data/slideshow/45545_Comment.pptx 100 1000"
  "pptx-open poi/test-data/slideshow/45545_Comment.pptx 50 500"
  "pptx-open-text poi/test-data/slideshow/45545_Comment.pptx 50 500"
  "pptx-open-text ooxml/pptx/SmartArt.pptx 100 1000"
)
: > "$OUT"
for round in $(seq 1 "$ROUNDS"); do
  for spec in "${SPECS[@]}"; do
    set -- $spec
    for leg in before after after before; do
      line=$(taskset -c "$CPU" "$S/bench-$leg" "$1" "$T/$2" "$3" "$4")
      echo "{\"round\":$round,\"leg\":\"$leg\",\"load1\":$(cut -d' ' -f1 /proc/loadavg),${line#\{}" >> "$OUT"
    done
  done
done
