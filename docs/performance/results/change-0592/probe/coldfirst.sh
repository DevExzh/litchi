#!/bin/bash
# Cost of the FIRST measured call in a fresh process, after the harness-style
# untimed text() preparation: whole-process cycles at reps=1 minus reps=0.
set -u
echo "op,order,leg,rep,cycles_first_call"
for order in 1 2 3 4; do
  case $order in 1|4) leg=before ;; 2|3) leg=after ;; esac
  bin=/home/zhuhe/code/litchi-worktrees/targets/0592-probe-$leg/release/probe0592
  for i in $(seq 1 25); do
    for op in eager_prepared_paragraph_count eager_paragraph_count; do
      a=$(taskset -c 12 perf stat -e cycles -x, $bin $op 200 0 2>&1 >/dev/null | awk -F, '/cycles/{print $1}')
      b=$(taskset -c 12 perf stat -e cycles -x, $bin $op 200 1 2>&1 >/dev/null | awk -F, '/cycles/{print $1}')
      echo "$op,$order,$leg,$i,$(( b - a ))"
    done
  done
done
