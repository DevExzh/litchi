#!/bin/bash
# Callgrind isolation pairs on the base binary (0654 method): samples 1 and 3,
# warmup 0; the difference is two measured iterations.
set -u
BASE=/home/zhuhe/code/litchi-worktrees/targets/base-009d515bef/release/litchi-perf-baseline
cd /home/zhuhe/code/litchi-worktrees/scratch/0747/cg
for shape in medium dense-sparse; do
  for s in 1 3; do
    taskset -c 24 valgrind --tool=callgrind --callgrind-out-file=cg-$shape-s$s.out \
      $BASE --case xlsx_source_backed_cell_values_one_edit_save --xlsx-cell-crud-shape $shape \
      --samples $s --warmup 0 --json run-$shape-s$s.json > log-$shape-s$s.txt 2>&1
    echo "$shape s$s exit $?" >> status.txt
  done
done
