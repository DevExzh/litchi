#!/bin/bash
# `perf stat` isolation pairs for change 0641: N and N+M samples, differenced and
# divided by M, five repetitions per leg, on one pinned CPU. The third leg is the
# before binary again, so the A/A floor for this metric is measured in the same
# window. Cycles are what callgrind cannot price (0604).
set -e
S=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0641
LOW=${LOW:-20}
HIGH=${HIGH:-120}
REPS=${REPS:-5}
CPU=24
cd /home/zhuhe/code/litchi-worktrees/0641
while IFS='|' read -r label file idx op; do
  [ -z "$label" ] && continue
  for rep in $(seq 1 $REPS); do
    for leg in before after before2; do
      bin=$leg; [ "$leg" = "before2" ] && bin=before
      for n in $LOW $HIGH; do
        taskset -c $CPU perf stat -x, -e cycles,instructions -r 5 \
          "$S/bin/xsa-$bin" --input "$file" --operation "$op" --worksheet-index "$idx" \
          --warmups 2 --samples "$n" > /dev/null 2> "$S/perf/$label-$leg-$n-$rep.txt"
      done
    done
  done
  echo "perf done $label"
done < "$1"
