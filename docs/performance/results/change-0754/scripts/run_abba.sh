#!/bin/bash
# Change 0754 before/after timing. Rounds of A B B A per case (A = base
# binary, B = branch binary, equal-length paths), every process pinned to
# core 12 and wrapped in `perf stat -e instructions,cycles`.
set -u
S=/home/zhuhe/code/litchi-worktrees/scratch/0754
OUT=${OUT:-$S/timing}
CORE=${CORE:-12}
ROUNDS=${ROUNDS:-6}
export TMPDIR=$S/tmp
mkdir -p "$OUT/raw"
# label|samples|warmup|harness arguments
CASES=(
  "noop_large|40|5|--case docx_semantic_noop_edit_save --semantic-shape large"
  "noop_medium|300|30|--case docx_semantic_noop_edit_save --semantic-shape medium"
  "one_large|40|5|--case docx_semantic_one_edit_save --semantic-shape large"
  "one_medium|300|30|--case docx_semantic_one_edit_save --semantic-shape medium"
  "pct_large|30|5|--case docx_semantic_one_percent_edit_save --semantic-shape large"
  "pct_medium|300|30|--case docx_semantic_one_percent_edit_save --semantic-shape medium"
  "text_large|40|5|--case docx_semantic_full_text --semantic-shape large"
  "text_medium|300|30|--case docx_semantic_full_text --semantic-shape medium"
  "sb_one|60|10|--case docx_source_backed_one_edit_save"
  "ordinary_lifecycle|60|10|--case docx_ordinary_save_lifecycle"
  "ctl_pptx_text|15|3|--case pptx_semantic_full_text --semantic-shape large"
  "ctl_xlsx_cell|300|30|--case xlsx_first_cell --xlsx-shape medium"
)
run() {
  local arm=$1 label=$2 samples=$3 warmup=$4 args=$5 round=$6 slot=$7
  local base="$OUT/raw/${label}-r${round}-s${slot}-${arm}"
  # shellcheck disable=SC2086
  taskset -c "$CORE" perf stat -x, -e instructions,cycles -o "$base.perf" \
    "$S/bin/lpb-$arm" $args --samples "$samples" --warmup "$warmup" --json "$base.json" \
    > "$base.log" 2>&1
  echo "$label round=$round slot=$slot arm=$arm exit=$?" >> "$OUT/status.txt"
}
for round in $(seq 1 "$ROUNDS"); do
  for entry in "${CASES[@]}"; do
    IFS='|' read -r label samples warmup args <<< "$entry"
    run A "$label" "$samples" "$warmup" "$args" "$round" 1
    run B "$label" "$samples" "$warmup" "$args" "$round" 2
    run B "$label" "$samples" "$warmup" "$args" "$round" 3
    run A "$label" "$samples" "$warmup" "$args" "$round" 4
  done
done
echo done >> "$OUT/status.txt"
