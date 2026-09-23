#!/bin/bash
# Callgrind isolation pairs (samples 1 and 3, warmup 0) for the self-built legs.
set -u
cd /home/zhuhe/code/litchi-worktrees/scratch/0747/cg
B=/home/zhuhe/code/litchi-worktrees/scratch/0747/bin
run() {
  local tag=$1 bin=$2 case=$3 shape=$4 s=$5
  taskset -c 24 valgrind --tool=callgrind --callgrind-out-file=cg-$tag-$shape-s$s.out \
    "$bin" --case "$case" --xlsx-cell-crud-shape "$shape" --samples "$s" --warmup 0 \
    --json run-$tag-$shape-s$s.json > log-$tag-$shape-s$s.txt 2>&1
  echo "$tag $case $shape s$s exit $?" >> status-final.txt
}
for shape in medium dense-sparse; do
  for s in 1 3; do
    run beforeself $B/litchi-perf-baseline.before-self xlsx_source_backed_cell_values_one_edit_save $shape $s
    run after2 $B/litchi-perf-baseline.after2 xlsx_source_backed_cell_values_one_edit_save $shape $s
  done
done
for s in 1 3; do
  run beforeself-eager $B/litchi-perf-baseline.before-self xlsx_eager_cell_values_one_edit_save dense-sparse $s
  run after2-eager $B/litchi-perf-baseline.after2 xlsx_eager_cell_values_one_edit_save dense-sparse $s
done
echo done >> status-final.txt
