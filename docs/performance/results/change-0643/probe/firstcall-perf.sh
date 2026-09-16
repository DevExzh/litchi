#!/bin/bash
# Part (a) counters. `perf stat` cannot be scoped to the timed region, so the
# whole process is counted at reps=0 and at reps=1 and the *medians* are
# differenced (a median of per-pair differences is dominated by process-start
# variance; a difference of medians is not). The pair therefore prices one
# first call in a fresh process, including what it pays for being first.
set -u
SP=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0643
BIN=/home/zhuhe/code/litchi-worktrees/targets/0643-bin
FIXTURE=${FIXTURE:-/home/zhuhe/code/litchi-worktrees/targets/0643-fixtures/harness-docx.docx}
LEGS="${LEGS:-rev0592 base}"
OPS="${OPS:-eager_prepared_paragraph_count eager_prepared_warm_paragraph_count eager_noprep_paragraph_count}"
REPEATS=${REPEATS:-40}
EV=cycles,instructions,branch-misses,dTLB-load-misses,page-faults
echo "op,leg,reps,rep,cycles,instructions,branch_misses,dtlb_load_misses,page_faults"
for op in $OPS; do
  for i in $(seq 1 $REPEATS); do
    for leg in $LEGS; do
      for reps in 0 1; do
        v=$(taskset -c 18 perf stat -e $EV -x, $BIN/probe-$leg $op "$FIXTURE" $reps 2>&1 >/dev/null \
          | awk -F, '/cycles/{c=$1} /instructions/{i=$1} /branch-misses/{b=$1} /dTLB-load-misses/{d=$1} /page-faults/{p=$1} END{print c","i","b","d","p}')
        echo "$op,$leg,$reps,$i,$v"
      done
    done
  done
done
