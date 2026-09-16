#!/usr/bin/env bash
# Change 0658: ABBA paired timing on CPU 13, plus an A/A floor in the same window.
#   (a) the ineligible source-backed one-cell read, through the retained probe;
#   (b) the harness's selected-cell and range cases, as the eligible controls.
set -u
S=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0658
STAGE=/home/zhuhe/code/litchi-worktrees/targets/0658-staged
CASES=xlsx_file_selected_cell,xlsx_range_source_first_cell,xlsx_narrow_column_range_scan
mkdir -p "$S/out/timing"

probe() { # probe <binary> <fixture> <tag>
  taskset -c 13 "$1" time "$S/fixtures/$2.xlsx" H680 5 30 20 \
    > "$S/out/timing/probe-$2-$3.txt" 2>"$S/out/timing/probe-$2-$3.err"
}
harness() { # harness <binary> <tag>
  taskset -c 13 "$1" --warmup 5 --samples 30 --case "$CASES" \
    --json "$S/out/timing/$2.json" >/dev/null 2>"$S/out/timing/$2.err"
}

for v in control real; do
  for leg in A1:probe-before B1:probe-after B2:probe-after A2:probe-before \
             F1:probe-before F2:probe-before F3:probe-before F4:probe-before; do
    tag=${leg%%:*}; bin=$STAGE/${leg##*:}
    probe "$bin" "$v" "$tag" || echo "FAILED probe $v $tag"
  done
done

for leg in A1:harness-before B1:harness-after B2:harness-after A2:harness-before \
           F1:harness-before F2:harness-before F3:harness-before F4:harness-before; do
  tag=${leg%%:*}; bin=$STAGE/${leg##*:}
  harness "$bin" "$tag" || echo "FAILED harness $tag"
done
echo timing-done
