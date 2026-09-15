#!/bin/bash
# Instruction isolation pairs. Per-operation Ir = (T_hi - T_lo) / (hi - lo).
set -u
SP=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0592
OUT=$SP/cg
mkdir -p $OUT
OPS="eager_document eager_text eager_write_text eager_paragraph_first eager_paragraph_repeat8 eager_paragraph_count eager_tables source_document source_text source_paragraph_first source_paragraph_count"
run() { # leg op shape reps
  local leg=$1 op=$2 shape=$3 reps=$4
  local bin=/home/zhuhe/code/litchi-worktrees/targets/0592-probe-$leg/release/probe0592
  local f=$OUT/$leg-$op-$shape-$reps.out
  taskset -c 12 valgrind --tool=callgrind --callgrind-out-file=$f \
    --cache-sim=no --branch-sim=no $bin $op $shape $reps > /dev/null 2>$OUT/$leg-$op-$shape-$reps.log
  grep -m1 '^summary:' $f | awk '{print $2}'
}
printf "%-26s %-6s %-7s %14s\n" operation shape leg ir_per_op
for shape in 200 10000; do
  if [ "$shape" = "200" ]; then lo=4; hi=20; else lo=1; hi=5; fi
  for op in $OPS; do
    for leg in before after; do
      a=$(run $leg $op $shape $lo)
      b=$(run $leg $op $shape $hi)
      printf "%-26s %-6s %-7s %14d\n" $op $shape $leg $(( (b - a) / (hi - lo) ))
    done
  done
done
