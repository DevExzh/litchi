#!/usr/bin/env bash
# change 0622: callgrind isolation pairs (N=1, N=11) per leg/case/shape.
set -uo pipefail
OUT=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0622/counts
mkdir -p "$OUT"
declare -A BIN=(
  [before]=/home/zhuhe/code/litchi-worktrees/targets/0622-before/release/litchi-perf-baseline
  [after]=/home/zhuhe/code/litchi-worktrees/0622/target-perf/release/litchi-perf-baseline
)
declare -A DIR=(
  [before]=/home/zhuhe/code/litchi-worktrees/before-1e4198321
  [after]=/home/zhuhe/code/litchi-worktrees/0622
)
jobs_running=0
for leg in before after; do
  for case in xlsx_source_backed_cell_values_one_edit_save xlsx_source_backed_cell_values_one_percent_edit_save; do
    for shape in medium dense-sparse noncompact; do
      if [ "$case" = xlsx_source_backed_cell_values_one_percent_edit_save ] && [ "$shape" = noncompact ]; then continue; fi
      for n in 1 11; do
        tag="$leg-${case#xlsx_source_backed_cell_values_}-$shape-n$n"
        (
          cd "${DIR[$leg]}" || exit 1
          taskset -c 15 valgrind --tool=callgrind --callgrind-out-file="$OUT/$tag.out" \
            --cache-sim=no --branch-sim=no \
            "${BIN[$leg]}" --case "$case" --xlsx-cell-crud-shape "$shape" \
            --warmup 0 --samples "$n" --json /dev/null > "$OUT/$tag.stdout" 2> "$OUT/$tag.stderr"
          echo "$tag exit=$?" >> "$OUT/progress.txt"
        ) &
        jobs_running=$((jobs_running+1))
        if [ "$jobs_running" -ge 3 ]; then wait -n; jobs_running=$((jobs_running-1)); fi
      done
    done
  done
done
wait
echo ALLDONE >> "$OUT/progress.txt"
