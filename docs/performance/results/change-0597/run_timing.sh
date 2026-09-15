#!/usr/bin/env bash
# Change 0597: ABBA paired timing on CPU 20, plus an A/A floor in the same window.
set -u
S=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0597
BEFORE=/home/zhuhe/code/litchi-worktrees/targets/0597-before/release/litchi-perf-baseline
AFTER=/home/zhuhe/code/litchi-worktrees/0597/target-perf/release/litchi-perf-baseline
CASES=xlsx_file_selected_cell,xlsx_range_source_first_cell,xlsx_narrow_column_range_scan
mkdir -p "$S/out/timing"
run() { # run <binary> <tag>
  taskset -c 20 "$1" --warmup 5 --samples 30 --case "$CASES" \
    --json "$S/out/timing/$2.json" >/dev/null 2>"$S/out/timing/$2.err"
}
for leg in A1:$BEFORE B1:$AFTER B2:$AFTER A2:$BEFORE; do
  tag=${leg%%:*}; bin=${leg##*:}
  run "$bin" "$tag" || echo "FAILED $tag"
done
# A/A floor in the same window
for leg in F1 F2 F3 F4; do run "$BEFORE" "$leg" || echo "FAILED $leg"; done
echo timing-done
