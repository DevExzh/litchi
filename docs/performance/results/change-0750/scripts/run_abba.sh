#!/bin/bash
# Change 0750 before/after timing (0747's method). Four rounds; in every round
# each case runs A B B A (A = base binary, B = candidate), every process
# pinned to one core. The harness is unchanged by the change.
set -u
BEFORE=${BEFORE:?}
AFTER=${AFTER:?}
OUT=${OUT:?}
CORE=${CORE:-24}
ROUNDS=${ROUNDS:-4}
mkdir -p "$OUT/raw"
# label|samples|warmup|harness arguments
CASES=(
  "sb_one_medium|30|5|--case xlsx_source_backed_cell_values_one_edit_save --xlsx-cell-crud-shape medium"
  "sb_one_dense|20|3|--case xlsx_source_backed_cell_values_one_edit_save --xlsx-cell-crud-shape dense-sparse"
  "eager_one_dense|20|3|--case xlsx_eager_cell_values_one_edit_save --xlsx-cell-crud-shape dense-sparse"
  "docx_sb_one|40|5|--case docx_source_backed_one_edit_save"
  "pptx_sb_one|40|5|--case pptx_source_backed_one_edit_save"
  "control_xls|20|3|--case xls_semantic_one_edit_save"
)
run() {
  local arm=$1 bin=$2 label=$3 samples=$4 warmup=$5 args=$6 round=$7 slot=$8
  local json="$OUT/raw/${label}-r${round}-s${slot}-${arm}.json"
  # shellcheck disable=SC2086
  taskset -c "$CORE" "$bin" $args --samples "$samples" --warmup "$warmup" --json "$json" \
    > "$OUT/raw/${label}-r${round}-s${slot}-${arm}.log" 2>&1
  echo "$label round=$round slot=$slot arm=$arm exit=$?" >> "$OUT/status.txt"
}
for round in $(seq 1 "$ROUNDS"); do
  for entry in "${CASES[@]}"; do
    IFS='|' read -r label samples warmup args <<< "$entry"
    run A "$BEFORE" "$label" "$samples" "$warmup" "$args" "$round" 1
    run B "$AFTER"  "$label" "$samples" "$warmup" "$args" "$round" 2
    run B "$AFTER"  "$label" "$samples" "$warmup" "$args" "$round" 3
    run A "$BEFORE" "$label" "$samples" "$warmup" "$args" "$round" 4
  done
done
echo done >> "$OUT/status.txt"
