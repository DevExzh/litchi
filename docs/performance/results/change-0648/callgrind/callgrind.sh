#!/bin/bash
# Callgrind isolation pairs for change 0648: N=2 and N=6 lifecycles of one
# operation, differenced and divided by 4.
set -e
S=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0648
CPU=23
LOW=${LOW:-2}
HIGH=${HIGH:-6}
cd /home/zhuhe/code/litchi-worktrees/0648
for spec in "$@"; do
  IFS='|' read -r kind op file label <<< "$spec"
  for leg in before after; do
    for n in $LOW $HIGH; do
      taskset -c $CPU valgrind --tool=callgrind --callgrind-out-file=/dev/null \
        "$S/bin/probe-$leg" time "$kind" "$op" "$n" "$file" \
        > /dev/null 2> "$S/cg/cg-$label-$leg-$n.log"
    done
  done
  echo "done $label"
done
