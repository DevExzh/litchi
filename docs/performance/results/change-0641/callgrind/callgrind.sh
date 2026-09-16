#!/bin/bash
# Callgrind isolation pairs for change 0641: profile N and N+M samples of one
# harness operation and difference the per-symbol totals, then divide by M.
set -e
S=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0641
LOW=2
HIGH=12
CPU=24
cd /home/zhuhe/code/litchi-worktrees/0641
while IFS='|' read -r label file idx op; do
  [ -z "$label" ] && continue
  for leg in before after; do
    for n in $LOW $HIGH; do
      taskset -c $CPU valgrind --tool=callgrind --callgrind-out-file="$S/cg/$label-$leg-$n.out" \
        "$S/bin/xsa-$leg" --input "$file" --operation "$op" --worksheet-index "$idx" \
        --warmups 1 --samples "$n" > /dev/null 2> "$S/cg/$label-$leg-$n.log"
    done
  done
  echo "callgrind done $label"
done < "$1"
