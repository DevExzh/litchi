#!/usr/bin/env bash
# change 0635: callgrind isolation pairs (N=1, N=11) per leg/case/shape.
set -uo pipefail
LEG="$1"
OUT=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0635/counts
mkdir -p "$OUT"
declare -A BIN=(
  [before]=/home/zhuhe/code/litchi-worktrees/targets/0635-staged/before-litchi-perf-baseline
  [after]=/home/zhuhe/code/litchi-worktrees/targets/0635-staged/after-litchi-perf-baseline
)
declare -A DIR=(
  [before]=/home/zhuhe/code/litchi-worktrees/before-c7326f680
  [after]=/home/zhuhe/code/litchi-worktrees/0635
)
run() { # tag, extra args...
  local tag="$1"; shift
  (
    cd "${DIR[$LEG]}" || exit 1
    taskset -c 11 valgrind --tool=callgrind --callgrind-out-file="$OUT/$tag.out" \
      --cache-sim=no --branch-sim=no \
      "${BIN[$LEG]}" "$@" --warmup 0 --json /dev/null > "$OUT/$tag.stdout" 2> "$OUT/$tag.stderr"
    echo "$tag exit=$?" >> "$OUT/progress.txt"
  )
}
jobs_running=0
sched() {
  run "$@" &
  jobs_running=$((jobs_running+1))
  if [ "$jobs_running" -ge 3 ]; then wait -n; jobs_running=$((jobs_running-1)); fi
}
for n in 1 11; do
  for shape in medium dense-sparse noncompact; do
    sched "$LEG-one_edit-$shape-n$n" --case xlsx_source_backed_cell_values_one_edit_save --xlsx-cell-crud-shape "$shape" --samples "$n"
  done
  for shape in medium dense-sparse; do
    sched "$LEG-one_percent-$shape-n$n" --case xlsx_source_backed_cell_values_one_percent_edit_save --xlsx-cell-crud-shape "$shape" --samples "$n"
  done
  sched "$LEG-managed_one_edit-medium-n$n" --case xlsx_source_backed_managed_cell_values_one_edit_save --xlsx-cell-crud-shape medium --samples "$n"
  sched "$LEG-producer_medium_edit-n$n" --case xlsx_producer_medium_source_one_edit_save --samples "$n"
  sched "$LEG-producer_dense_edit-n$n" --case xlsx_producer_dense_source_one_edit_save --samples "$n"
done
wait
echo "ALLDONE-$LEG" >> "$OUT/progress.txt"
