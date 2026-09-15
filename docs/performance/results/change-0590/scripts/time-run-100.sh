#!/bin/bash
# usage: time-run.sh <case> <samples> <warmup>
set -u
S=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0590
B=/home/zhuhe/code/litchi-worktrees/targets/0590-before/release/litchi-perf-baseline
A=/home/zhuhe/code/litchi-worktrees/0590/tools/perf-baseline/target/release/litchi-perf-baseline
case=$1; samples=$2; warmup=$3
for leg in A1:$B B1:$A B2:$A A2:$B; do
  name=${leg%%:*}; bin=${leg##*:}
  out="$S/time100-$case-$name.json"
  rm -f "$out"
  taskset -c 10 "$bin" --warmup "$warmup" --samples "$samples" --case "$case" --json "$out" > "$S/time100-$case-$name.log" 2>&1
  echo "$name exit=$?"
done
