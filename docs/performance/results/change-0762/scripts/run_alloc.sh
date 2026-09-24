#!/usr/bin/env bash
# Allocation counts (changes 0762/0763): one process per leg and case, CPU 12.
set -euo pipefail
OUT=$1
BIN=/home/zhuhe/code/litchi-worktrees/scratch/0762/bin
export TMPDIR=/home/zhuhe/code/litchi-worktrees/scratch/0762/tmp
mkdir -p "$OUT"
CASES=(
  "streaming|--case docx_streaming_create,xlsx_streaming_create,pptx_streaming_create --semantic-shape tiny,medium,large --samples 3 --warmup 1"
  "ctl-docx-semantic-one-edit|--case docx_semantic_one_edit_save --semantic-shape large --samples 5 --warmup 1"
  "ctl-xlsx-ordinary-save|--case xlsx_ordinary_save_lifecycle --samples 5 --warmup 1"
)
for entry in "${CASES[@]}"; do
  name=${entry%%|*}
  args=${entry#*|}
  for leg in before after; do
    bin=$BIN/b/litchi-perf-baseline-alloc
    [ "$leg" = after ] && bin=$BIN/a/litchi-perf-baseline-alloc
    # shellcheck disable=SC2086
    taskset -c 12 "$bin" $args --json "$OUT/$name-$leg.json" > /dev/null 2>> "$OUT/stderr.log"
  done
done
echo ALLOC-COMPLETE >> "$OUT/stderr.log"
