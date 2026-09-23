#!/usr/bin/env bash
# ABBA timing campaign for change 0752.
# Usage: run_abba.sh OUTDIR ROUNDS
# Every process is pinned to CPU 28. In each round every case runs
# before, after, after, before. Raw harness JSON reports are kept.
set -euo pipefail
OUT=$1
ROUNDS=${2:-4}
BEFORE=/home/zhuhe/code/litchi-worktrees/targets/0752-before/release/litchi-perf-baseline
AFTER=/home/zhuhe/code/litchi-worktrees/targets/0752/release/litchi-perf-baseline
export TMPDIR=/home/zhuhe/code/litchi-worktrees/scratch/0752/tmp
mkdir -p "$OUT"
# name|harness arguments
CASES=(
  "docx-large|--case docx_streaming_create --semantic-shape large --samples 20 --warmup 3"
  "docx-medium|--case docx_streaming_create --semantic-shape medium --samples 40 --warmup 5"
  "xlsx-large|--case xlsx_streaming_create --semantic-shape large --samples 15 --warmup 3"
  "xlsx-medium|--case xlsx_streaming_create --semantic-shape medium --samples 30 --warmup 5"
  "pptx-large|--case pptx_streaming_create --semantic-shape large --samples 15 --warmup 3"
  "pptx-medium|--case pptx_streaming_create --semantic-shape medium --samples 30 --warmup 5"
  "ctl-docx-semantic-one-edit|--case docx_semantic_one_edit_save --semantic-shape large --samples 40 --warmup 5"
  "ctl-xlsx-ordinary-save|--case xlsx_ordinary_save_lifecycle --samples 20 --warmup 3"
  "ctl-cfb-shared-bulk|--case cfb_open_stream_mini_shared_bulk --samples 40 --warmup 5"
)
for round in $(seq 1 "$ROUNDS"); do
  for entry in "${CASES[@]}"; do
    name=${entry%%|*}
    args=${entry#*|}
    slot=0
    for leg in before after after before; do
      slot=$((slot + 1))
      bin=$BEFORE
      [ "$leg" = after ] && bin=$AFTER
      json="$OUT/$name-r$round-s$slot-$leg.json"
      # shellcheck disable=SC2086
      taskset -c 28 "$bin" $args --json "$json" > /dev/null 2>> "$OUT/stderr.log"
      echo "$(date +%T) round $round $name slot $slot $leg done" >> "$OUT/progress.log"
    done
  done
done
echo "CAMPAIGN-COMPLETE" >> "$OUT/progress.log"
