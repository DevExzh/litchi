#!/bin/bash
# Native cycle isolation pairs on CPU 12, order before/after/after/before.
set -u
SP=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0592
OPS="eager_document eager_text eager_paragraph_count eager_paragraph_first"
SHAPE=${1:-200}
LO=4; HI=20
cyc() { # leg op reps
  local bin=/home/zhuhe/code/litchi-worktrees/targets/0592-probe-$1/release/probe0592
  taskset -c 12 perf stat -e cycles -x, $bin $2 $SHAPE $3 2>&1 >/dev/null | awk -F, '/cycles/{print $1}'
}
echo "op,shape,order,leg,rep,cycles_per_op"
for op in $OPS; do
  for order in 1 2 3 4; do
    case $order in 1|4) leg=before ;; 2|3) leg=after ;; esac
    for i in $(seq 1 15); do
      a=$(cyc $leg $op $LO); b=$(cyc $leg $op $HI)
      echo "$op,$SHAPE,$order,$leg,$i,$(( (b - a) / (HI - LO) ))"
    done
  done
done
