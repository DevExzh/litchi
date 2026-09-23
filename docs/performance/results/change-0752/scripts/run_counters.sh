#!/usr/bin/env bash
# Per-iteration user-space instructions and cycles for change 0752, by
# differencing a LOW-sample and a HIGH-sample run of the same binary and case
# (perf_delta.py). Two rounds of before, after, after, before; CPU 28.
set -euo pipefail
OUT=$1
BEFORE=/home/zhuhe/code/litchi-worktrees/targets/0752-before/release/litchi-perf-baseline
AFTER=/home/zhuhe/code/litchi-worktrees/targets/0752/release/litchi-perf-baseline
DELTA=/home/zhuhe/code/litchi-worktrees/scratch/0752/scripts/perf_delta.py
TMP=/home/zhuhe/code/litchi-worktrees/scratch/0752/tmp
mkdir -p "$OUT"
# name|case|semantic shape (or -)|low|high|warmup
CASES=(
  "docx-large|docx_streaming_create|large|2|12|2"
  "docx-medium|docx_streaming_create|medium|5|45|3"
  "xlsx-large|xlsx_streaming_create|large|2|10|2"
  "xlsx-medium|xlsx_streaming_create|medium|5|45|3"
  "pptx-large|pptx_streaming_create|large|2|10|2"
  "pptx-medium|pptx_streaming_create|medium|5|45|3"
  "ctl-docx-semantic-one-edit|docx_semantic_one_edit_save|large|5|45|3"
  "ctl-xlsx-ordinary-save|xlsx_ordinary_save_lifecycle|-|5|45|3"
  "ctl-cfb-shared-bulk|cfb_open_stream_mini_shared_bulk|-|5|85|3"
)
for round in 1 2; do
  for entry in "${CASES[@]}"; do
    IFS='|' read -r name case shape low high warmup <<< "$entry"
    [ "$shape" = "-" ] && shape=""
    slot=0
    for leg in before after after before; do
      slot=$((slot + 1))
      bin=$BEFORE
      [ "$leg" = after ] && bin=$AFTER
      target="$OUT/$name-r$round-s$slot-$leg.json"
      # Resume: a completed measurement is kept, an empty one is redone.
      [ -s "$target" ] && continue
      python3 "$DELTA" "$bin" "$case" "$shape" "$low" "$high" "$warmup" 28 "$TMP" > "$target"
      echo "$(date +%T) round $round $name slot $slot $leg done" >> "$OUT/progress.log"
    done
  done
done
echo "COUNTERS-COMPLETE" >> "$OUT/progress.log"
