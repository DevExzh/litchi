#!/bin/bash
# usage: time-run.sh <case> <samples> <warmup>
set -u
S=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0645
B=$S/stage/before
A=$S/stage/afterb
case=$1; samples=$2; warmup=$3
for leg in A1:$B B1:$A B2:$A A2:$B; do
  name=${leg%%:*}; bin=${leg##*:}
  out="$S/out/time-$case-$name.json"
  rm -f "$out"
  taskset -c 20 "$bin" --warmup "$warmup" --samples "$samples" --case "$case" --json "$out" > "$S/out/time-$case-$name.log" 2>&1
  echo "$name exit=$?"
done
