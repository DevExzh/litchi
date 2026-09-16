#!/bin/bash
# Part (a): the first `document() + paragraph_count()` call in a fresh process,
# timed in-process exactly as the perf-baseline filesystem child times it
# (`Instant` around the call, one call per process).  One process per sample.
#
# Three operations differ only in the untimed preparation:
#   eager_prepared_paragraph_count       prep = document().text()   (the harness)
#   eager_prepared_warm_paragraph_count  prep = text() + paragraph_count()
#   eager_noprep_paragraph_count         no preparation
# Order A1 B1 B2 A2, A = 0592 reverted (eager index), B = base (lazy index).
set -u
SP=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0643
BIN=/home/zhuhe/code/litchi-worktrees/targets/0643-bin
OUT=$SP/out/firstcall
FIXTURE=${FIXTURE:-/home/zhuhe/code/litchi-worktrees/targets/0643-fixtures/harness-docx.docx}
TAG=${TAG:-harness}
SAMPLES=${SAMPLES:-60}
mkdir -p $OUT
for op in eager_prepared_paragraph_count eager_prepared_warm_paragraph_count eager_noprep_paragraph_count; do
  for run in A1 B1 B2 A2; do
    case $run in A*) leg=rev0592 ;; B*) leg=base ;; esac
    : > $OUT/$op-$TAG-$run.txt
    for i in $(seq 1 $SAMPLES); do
      taskset -c 18 $BIN/probe-$leg timed $op "$FIXTURE" 1 >> $OUT/$op-$TAG-$run.txt
    done
  done
done
