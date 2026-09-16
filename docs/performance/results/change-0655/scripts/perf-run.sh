#!/bin/bash
# usage: perf-run.sh <case> <samples> <warmup>   (A1 B1 B2 A2 order, cycles+instructions)
set -u
S=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0655
B=$S/stage/before
A=$S/stage/after
case=$1; samples=$2; warmup=$3
DECK=/home/zhuhe/code/litchi-worktrees/0655/test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-section-test.pptx
EXTRA=()
case "$case" in *real_file*) EXTRA=(--ooxml-file "$DECK");; esac
for leg in A1:$B B1:$A B2:$A A2:$B; do
  name=${leg%%:*}; bin=${leg##*:}
  rm -f "$S/out/perf-$case-$name.json"
  perf stat -e cycles,instructions,task-clock -x, -o "$S/out/perf-$case-$name.txt" -- \
    taskset -c 10 "$bin" --warmup "$warmup" --samples "$samples" --case "$case" "${EXTRA[@]}" \
    --json "$S/out/perf-$case-$name.json" > /dev/null 2>&1
  echo "$name exit=$? $(grep -E '^[0-9]' "$S/out/perf-$case-$name.txt" | cut -d, -f1,3 | tr '\n' ' ')"
done
