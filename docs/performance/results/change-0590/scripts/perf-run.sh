#!/bin/bash
# usage: perf-run.sh <case> <samples> <warmup>   (A1 B1 B2 A2 order, cycles+instructions)
set -u
S=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0590
B=/home/zhuhe/code/litchi-worktrees/targets/0590-before/release/litchi-perf-baseline
A=/home/zhuhe/code/litchi-worktrees/0590/tools/perf-baseline/target/release/litchi-perf-baseline
case=$1; samples=$2; warmup=$3
for leg in A1:$B B1:$A B2:$A A2:$B; do
  name=${leg%%:*}; bin=${leg##*:}
  rm -f "$S/perf-$case-$name.json"
  perf stat -e cycles,instructions,task-clock -x, -o "$S/perf-$case-$name.txt" -- \
    taskset -c 10 "$bin" --warmup "$warmup" --samples "$samples" --case "$case" \
    --json "$S/perf-$case-$name.json" > /dev/null 2>&1
  echo "$name exit=$? $(grep -E '^[0-9]' "$S/perf-$case-$name.txt" | cut -d, -f1,3 | tr '\n' ' ')"
done
