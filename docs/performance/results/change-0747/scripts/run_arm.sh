#!/bin/bash
# Callgrind isolation pairs (0654 method): samples 1 and 3, warmup 0, for one
# binary; the difference is two measured iterations.
set -u
BIN=${BIN:?}
TAG=${TAG:?}
cd /home/zhuhe/code/litchi-worktrees/scratch/0747/cg
for shape in medium dense-sparse; do
  for s in 1 3; do
    taskset -c 24 valgrind --tool=callgrind --callgrind-out-file=cg-$TAG-$shape-s$s.out \
      "$BIN" --case xlsx_source_backed_cell_values_one_edit_save --xlsx-cell-crud-shape $shape \
      --samples $s --warmup 0 --json run-$TAG-$shape-s$s.json > log-$TAG-$shape-s$s.txt 2>&1
    echo "$TAG $shape s$s exit $?" >> status-$TAG.txt
  done
done
