#!/bin/bash
# Allocation isolation pairs: reps=4 and reps=20, per-operation = (T20 - T4) / 16.
# Deterministic: the counting global allocator in the probe is not timing
# dependent, so these runs are not pinned and do not contend with the timed
# legs.
set -u
SP=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0643
OPS="eager_document eager_text eager_write_text source_text source_write_text"
LEGS="${LEGS:-base after}"
printf "%-22s %-8s %-8s %14s %16s\n" operation shape leg allocations allocated_bytes
for shape in 24 200 10000; do
  for op in $OPS; do
    for leg in $LEGS; do
      BIN=/home/zhuhe/code/litchi-worktrees/targets/0643-probe-$leg/release/probe0643
      lo=$($BIN $op $shape 4  | sed 's/.*allocations=\([0-9]*\) allocated_bytes=\([0-9]*\)/\1 \2/')
      hi=$($BIN $op $shape 20 | sed 's/.*allocations=\([0-9]*\) allocated_bytes=\([0-9]*\)/\1 \2/')
      a=$(( ( $(echo $hi | cut -d' ' -f1) - $(echo $lo | cut -d' ' -f1) ) / 16 ))
      b=$(( ( $(echo $hi | cut -d' ' -f2) - $(echo $lo | cut -d' ' -f2) ) / 16 ))
      printf "%-22s %-8s %-8s %14d %16d\n" $op $shape $leg $a $b
    done
  done
done
