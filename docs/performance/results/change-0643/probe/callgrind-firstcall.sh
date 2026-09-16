#!/bin/bash
# Part (a): callgrind over the selector's timed region. reps=0 and reps=1 are
# differenced, so the figure is the one `document() + paragraph_count()` call
# the harness times. Three legs: 0592 reverted, the layout control, and base.
set -u
SP=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0643
BIN=/home/zhuhe/code/litchi-worktrees/targets/0643-bin
F=/home/zhuhe/code/litchi-worktrees/targets/0643-fixtures/harness-docx.docx
OUT=$SP/cg
mkdir -p $OUT
printf "%-38s %-10s %10s %14s %14s %14s\n" operation leg reps total_ir call_ir ""
for op in eager_prepared_paragraph_count eager_noprep_paragraph_count; do
  for leg in rev0592 layoutctl base; do
    declare -A T
    for reps in 0 1; do
      f=$OUT/fc-$leg-$op-$reps.out
      taskset -c 19 valgrind --tool=callgrind --callgrind-out-file=$f --cache-sim=no --branch-sim=no \
        $BIN/probe-$leg $op "$F" $reps > /dev/null 2>$f.log
      T[$reps]=$(grep -m1 '^summary:' $f | awk '{print $2}')
    done
    printf "%-38s %-10s %10s %14d %14d\n" $op $leg "0/1" ${T[0]} $(( ${T[1]} - ${T[0]} ))
  done
done
