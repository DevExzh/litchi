#!/bin/bash
# Part (b) paired timing, order A1 B1 B2 A2 on CPU 18.  A = base (the owning
# sink parser), B = this branch (the borrowing one).  A1 against A2 is the
# floor.  One process per leg per case: the preparation runs once, then the
# operation is timed `SAMPLES` times.
set -u
SP=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0643
OUT=$SP/out/timing-sink
SAMPLES=${SAMPLES:-60}
mkdir -p $OUT
for run in A1 B1 B2 A2; do
  case $run in A*) leg=base ;; B*) leg=after ;; esac
  BIN=/home/zhuhe/code/litchi-worktrees/targets/0643-bin/probe-$leg
  for shape in 24 200 10000; do
    for op in eager_write_text source_write_text eager_text; do
      taskset -c 18 $BIN timed $op $shape $SAMPLES > $OUT/$op-$shape-$run.txt 2>/dev/null
    done
  done
  # the harness's own 200-paragraph corpus file, through PreparedDocx::eager
  for op in eager_write_text eager_text; do
    taskset -c 18 $BIN timed $op /home/zhuhe/code/litchi-worktrees/targets/0643-fixtures/harness-docx.docx $SAMPLES \
      > $OUT/$op-harness-$run.txt 2>/dev/null
  done
  echo "$run done"
done
