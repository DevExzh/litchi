#!/usr/bin/env bash
# change 0635: paired A1 B1 B2 A2 timing plus a four-run A/A floor, CPU 11.
# Two selector families per tag: the 0538/0550/0622 source-backed cell-value
# cases over three shapes, and the 0601 producer-shaped edit selectors.
set -euo pipefail
STAGE=/home/zhuhe/code/litchi-worktrees/targets/0635-staged
BEFORE="$STAGE/before-litchi-perf-baseline"
AFTER="$STAGE/after-litchi-perf-baseline"
OUT="$1"
CASES=xlsx_source_backed_cell_values_one_edit_save,xlsx_source_backed_cell_values_one_percent_edit_save,xlsx_source_backed_cell_values_batch_edit_save,xlsx_source_backed_managed_cell_values_one_edit_save,xlsx_source_backed_managed_cell_values_one_percent_edit_save
SHAPES=medium,dense-sparse,noncompact
PRODUCER=xlsx_producer_medium_source_one_edit_save,xlsx_producer_dense_source_one_edit_save
mkdir -p "$OUT"
run() {
  local bin="$1" tag="$2" dir="$3"
  ( cd "$dir" && taskset -c 11 "$bin" --case "$CASES" --xlsx-cell-crud-shape "$SHAPES" \
      --warmup 5 --samples 30 --json "$OUT/$tag.json" >/dev/null 2>"$OUT/$tag.stderr" )
  ( cd "$dir" && taskset -c 11 "$bin" --case "$PRODUCER" \
      --warmup 5 --samples 30 --json "$OUT/$tag-producer.json" >/dev/null 2>"$OUT/$tag-producer.stderr" )
  echo "$tag done"
}
BDIR=/home/zhuhe/code/litchi-worktrees/before-c7326f680
ADIR=/home/zhuhe/code/litchi-worktrees/0635
run "$BEFORE" A1 "$BDIR"
run "$AFTER"  B1 "$ADIR"
run "$AFTER"  B2 "$ADIR"
run "$BEFORE" A2 "$BDIR"
run "$BEFORE" F1 "$BDIR"
run "$BEFORE" F2 "$BDIR"
run "$BEFORE" F3 "$BDIR"
run "$BEFORE" F4 "$BDIR"
