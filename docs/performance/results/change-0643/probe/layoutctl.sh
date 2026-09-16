#!/bin/bash
# Part (a) layout control (0623's method). Three binaries, interleaved
# A1 C1 B1 B2 C2 A2 so each leg carries its own floor:
#   A  rev0592    no `OnceLock` at all; the scan runs inside `from_part`.
#   C  layoutctl  the `OnceLock` field and the `get_or_init` initializer are
#                 linked exactly as on base, but `from_part` fills the cell, so
#                 the lazy path is never taken and the scan runs where A runs it.
#   B  base       the `OnceLock` filled by the first paragraph query.
# C against A isolates the field and the code layout; B against C isolates the
# placement of the scan.
set -u
SP=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0643
BIN=/home/zhuhe/code/litchi-worktrees/targets/0643-bin
OUT=$SP/out/layoutctl
FIXTURE=${FIXTURE:-/home/zhuhe/code/litchi-worktrees/targets/0643-fixtures/harness-docx.docx}
TAG=${TAG:-harness}
SAMPLES=${SAMPLES:-60}
mkdir -p $OUT
for op in eager_prepared_paragraph_count eager_prepared_warm_paragraph_count eager_noprep_paragraph_count; do
  for run in A1 C1 B1 B2 C2 A2; do
    case $run in A*) leg=rev0592 ;; B*) leg=base ;; C*) leg=layoutctl ;; esac
    : > $OUT/$op-$TAG-$run.txt
    for i in $(seq 1 $SAMPLES); do
      taskset -c 18 $BIN/probe-$leg timed $op "$FIXTURE" 1 >> $OUT/$op-$TAG-$run.txt
    done
  done
done
