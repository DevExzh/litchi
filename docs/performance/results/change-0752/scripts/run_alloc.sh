#!/usr/bin/env bash
# Allocation counts for change 0752: one process per leg and case, CPU 28.
set -euo pipefail
OUT=$1
BEFORE=/home/zhuhe/code/litchi-worktrees/targets/0752-before/release/litchi-perf-baseline-alloc
AFTER=/home/zhuhe/code/litchi-worktrees/targets/0752/release/litchi-perf-baseline-alloc
export TMPDIR=/home/zhuhe/code/litchi-worktrees/scratch/0752/tmp
mkdir -p "$OUT"
CASES=(
  "streaming|--case docx_streaming_create,xlsx_streaming_create,pptx_streaming_create --semantic-shape tiny,medium,large --samples 3 --warmup 1"
  "ctl-docx-semantic-one-edit|--case docx_semantic_one_edit_save --semantic-shape large --samples 5 --warmup 1"
  "ctl-xlsx-ordinary-save|--case xlsx_ordinary_save_lifecycle --samples 5 --warmup 1"
  "ctl-cfb-shared-bulk|--case cfb_open_stream_mini_shared_bulk --samples 5 --warmup 1"
)
for entry in "${CASES[@]}"; do
  name=${entry%%|*}
  args=${entry#*|}
  for leg in before after; do
    bin=$BEFORE
    [ "$leg" = after ] && bin=$AFTER
    # shellcheck disable=SC2086
    taskset -c 28 "$bin" $args --json "$OUT/$name-$leg.json" > /dev/null 2>> "$OUT/stderr.log"
  done
done
