#!/usr/bin/env bash
# ABBA timing campaign (changes 0762/0763).
# Usage: run_abba.sh OUTDIR ROUNDS
# Every process is pinned to CPU 12. In each round every case runs before,
# after, after, before. The two legs' binaries are copies at paths of equal
# length (bin/b and bin/a). Raw harness JSON reports are kept.
set -euo pipefail
OUT=$1
ROUNDS=${2:-4}
BIN=/home/zhuhe/code/litchi-worktrees/scratch/0762/bin
export TMPDIR=/home/zhuhe/code/litchi-worktrees/scratch/0762/tmp
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
)
for round in $(seq 1 "$ROUNDS"); do
  for entry in "${CASES[@]}"; do
    name=${entry%%|*}
    args=${entry#*|}
    slot=0
    for leg in before after after before; do
      slot=$((slot + 1))
      bin=$BIN/b/litchi-perf-baseline
      [ "$leg" = after ] && bin=$BIN/a/litchi-perf-baseline
      json="$OUT/$name-r$round-s$slot-$leg.json"
      [ -s "$json" ] && continue
      # shellcheck disable=SC2086
      taskset -c 12 "$bin" $args --json "$json" > /dev/null 2>> "$OUT/stderr.log"
      echo "$(date +%T) round $round $name slot $slot $leg done" >> "$OUT/progress.log"
    done
  done
done
echo "CAMPAIGN-COMPLETE" >> "$OUT/progress.log"
