#!/bin/bash
# perf stat isolation pairs for change 0636: N and N+M iterations, differenced.
set -e
S=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0636
LOW=${LOW:-20}
HIGH=${HIGH:-120}
CPU=12
cd /home/zhuhe/code/litchi-worktrees/0636
for spec in "$@"; do
  IFS='|' read -r kind op file label <<< "$spec"
  for leg in before after; do
    for n in $LOW $HIGH; do
      taskset -c $CPU perf stat -x, -e cycles,instructions -r 5 \
        "$S/bin/probe-$leg" time "$kind" "$op" "$n" "$file" \
        > /dev/null 2> "$S/cg/perf-$label-$leg-$n.txt"
    done
  done
done
