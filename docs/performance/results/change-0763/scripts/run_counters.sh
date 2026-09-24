#!/usr/bin/env bash
# Per-iteration user-space instructions and cycles (changes 0762/0763), by
# differencing a LOW-sample and a HIGH-sample run of the same binary and case
# (perf_delta.py). Two rounds of before, after, after, before; CPU 12.
set -euo pipefail
OUT=$1
BIN=/home/zhuhe/code/litchi-worktrees/scratch/0762/bin
DELTA=/home/zhuhe/code/litchi-worktrees/scratch/0762/scripts/perf_delta.py
TMP=/home/zhuhe/code/litchi-worktrees/scratch/0762/tmp
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
)
for round in 1 2; do
  for entry in "${CASES[@]}"; do
    IFS='|' read -r name case shape low high warmup <<< "$entry"
    [ "$shape" = "-" ] && shape=""
    slot=0
    for leg in before after after before; do
      slot=$((slot + 1))
      bin=$BIN/b/litchi-perf-baseline
      [ "$leg" = after ] && bin=$BIN/a/litchi-perf-baseline
      target="$OUT/$name-r$round-s$slot-$leg.json"
      [ -s "$target" ] && continue
      python3 "$DELTA" "$bin" "$case" "$shape" "$low" "$high" "$warmup" 12 "$TMP" > "$target"
      echo "$(date +%T) round $round $name slot $slot $leg done" >> "$OUT/progress.log"
    done
  done
done
echo "COUNTERS-COMPLETE" >> "$OUT/progress.log"
