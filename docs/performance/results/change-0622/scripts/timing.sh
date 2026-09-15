#!/usr/bin/env bash
# change 0622: paired A1 B1 B2 A2 timing plus a four-run A/A floor, CPU 15.
set -euo pipefail
BEFORE=/home/zhuhe/code/litchi-worktrees/targets/0622-before/release/litchi-perf-baseline
AFTER=/home/zhuhe/code/litchi-worktrees/0622/target-perf/release/litchi-perf-baseline
OUT="$1"
CASES=xlsx_source_backed_cell_values_one_edit_save,xlsx_source_backed_cell_values_one_percent_edit_save,xlsx_source_backed_cell_values_batch_edit_save,xlsx_source_backed_managed_cell_values_one_edit_save,xlsx_source_backed_managed_cell_values_one_percent_edit_save
SHAPES=medium,dense-sparse,noncompact
mkdir -p "$OUT"
run() {
  local bin="$1" tag="$2" dir="$3"
  ( cd "$dir" && taskset -c 15 "$bin" --case "$CASES" --xlsx-cell-crud-shape "$SHAPES" \
      --warmup 5 --samples 30 --json "$OUT/$tag.json" >/dev/null 2>"$OUT/$tag.stderr" )
  echo "$tag done"
}
run "$BEFORE" A1 /home/zhuhe/code/litchi-worktrees/before-1e4198321
run "$AFTER"  B1 /home/zhuhe/code/litchi-worktrees/0622
run "$AFTER"  B2 /home/zhuhe/code/litchi-worktrees/0622
run "$BEFORE" A2 /home/zhuhe/code/litchi-worktrees/before-1e4198321
run "$BEFORE" F1 /home/zhuhe/code/litchi-worktrees/before-1e4198321
run "$BEFORE" F2 /home/zhuhe/code/litchi-worktrees/before-1e4198321
run "$BEFORE" F3 /home/zhuhe/code/litchi-worktrees/before-1e4198321
run "$BEFORE" F4 /home/zhuhe/code/litchi-worktrees/before-1e4198321
