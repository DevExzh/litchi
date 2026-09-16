#!/bin/bash
# Instruction isolation pairs. Per-operation Ir = (T_hi - T_lo) / (hi - lo).
set -u
SP=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0643
BIN=/home/zhuhe/code/litchi-worktrees/targets/0643-bin
OUT=$SP/cg
mkdir -p $OUT
run() { # leg op fixture reps -> total Ir
  local f=$OUT/$1-$2-$(basename $3)-$4.out
  taskset -c 19 valgrind --tool=callgrind --callgrind-out-file=$f \
    --cache-sim=no --branch-sim=no $BIN/probe-$1 $2 "$3" $4 > /dev/null 2>$f.log
  grep -m1 '^summary:' $f | awk '{print $2}'
}
printf "%-22s %-8s %-10s %14s\n" operation shape leg ir_per_op
for shape in 24 200 10000; do
  if [ "$shape" = "10000" ]; then lo=1; hi=5; else lo=4; hi=20; fi
  for op in eager_document eager_text eager_write_text source_text source_write_text; do
    for leg in base after; do
      a=$(run $leg $op $shape $lo); b=$(run $leg $op $shape $hi)
      printf "%-22s %-8s %-10s %14d\n" $op $shape $leg $(( (b - a) / (hi - lo) ))
    done
  done
done
