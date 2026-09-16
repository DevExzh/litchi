#!/usr/bin/env bash
# change 0635: change-0538 phase allocation regions, both legs, CPU 11.
set -uo pipefail
STAGE=/home/zhuhe/code/litchi-worktrees/targets/0635-staged
OUT="${1:?usage: alloc.sh OUTDIR}"
mkdir -p "$OUT"
for leg in before after; do
  bin="$STAGE/$leg-litchi-perf-baseline-alloc"
  dir=/home/zhuhe/code/litchi-worktrees/before-c7326f680
  [ "$leg" = after ] && dir=/home/zhuhe/code/litchi-worktrees/0635
  for case in xlsx_source_backed_cell_values_one_edit_save xlsx_source_backed_cell_values_one_percent_edit_save; do
    for shape in medium dense-sparse noncompact; do
      ( cd "$dir" && taskset -c 11 "$bin" --case "$case" --xlsx-cell-crud-shape "$shape" \
          --warmup 0 --samples 5 --json "$OUT/alloc-$leg-$case-$shape.json" >/dev/null 2>&1 )
      echo "$leg $case $shape done"
    done
  done
done
