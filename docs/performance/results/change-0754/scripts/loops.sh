#!/bin/bash
# Change 0754: per-iteration instructions, cycles and allocations of the timed
# operations, from the probe's `loop` mode at two iteration counts (the
# difference cancels the process's fixed work), each leg pinned to core 12.
set -u
PB=$1; PA=$2; OUT=$3
S=/home/zhuhe/code/litchi-worktrees/scratch/0754
for spec in "scan $S/corpus/semantic-large-document.xml" "text $S/corpus/semantic-large.docx" "noop $S/corpus/semantic-large.docx" "one $S/corpus/semantic-large.docx" "scan $S/corpus/semantic-medium-document.xml" "text $S/corpus/semantic-medium.docx" "noop $S/corpus/semantic-medium.docx" "one $S/corpus/semantic-medium.docx"; do
  for leg in A B; do
    bin=$PB; [ $leg = B ] && bin=$PA
    for n in 10 60; do
      taskset -c 12 perf stat -x, -e instructions,cycles -o $OUT/$leg-$(echo $spec | awk '{print $1}')-$(basename $(echo $spec | awk '{print $2}'))-$n.perf $bin loop $spec $n 2> $OUT/$leg-$(echo $spec | awk '{print $1}')-$(basename $(echo $spec | awk '{print $2}'))-$n.log
    done
  done
done
