#!/bin/bash
# usage: time-run.sh <case> <samples> <warmup>
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
  out="$S/out/time-$case-$name.json"
  rm -f "$out"
  taskset -c 10 "$bin" --warmup "$warmup" --samples "$samples" --case "$case" "${EXTRA[@]}" --json "$out" > "$S/out/time-$case-$name.log" 2>&1
  echo "$name exit=$?"
done
