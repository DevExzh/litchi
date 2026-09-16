#!/bin/bash
# Callgrind isolation pairs for change 0636: profile N and N+M iterations of one
# operation and difference the totals, then divide by M.
set -e
S=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0636
LOW=2
HIGH=6
cd /home/zhuhe/code/litchi-worktrees/0636
for leg in before after; do
  for spec in "$@"; do
    IFS='|' read -r kind op file label <<< "$spec"
    for n in $LOW $HIGH; do
      valgrind --tool=callgrind --callgrind-out-file="$S/cg/$label-$leg-$n.out" \
        "$S/bin/probe-$leg" time "$kind" "$op" "$n" "$file" > /dev/null 2> "$S/cg/$label-$leg-$n.log"
    done
  done
done
