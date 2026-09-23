#!/usr/bin/env bash
# Change 0744 allocator lane: separate allocator-metrics binaries, A B B A,
# CPU 12, commit/save cases only (the only XLSX cases that carry operation
# allocation metrics). Elapsed times from these processes are not used.
set -euo pipefail
A=${A:-/home/zhuhe/code/litchi-worktrees/targets/0744-before/release/litchi-perf-baseline-alloc}
B=${B:-/home/zhuhe/code/litchi-worktrees/targets/0744/release/litchi-perf-baseline-alloc}
OUT=${OUT:-alloc}
mkdir -p "$OUT"
for slot in A1 B1 B2 A2; do
  bin=$A; [ "${slot:0:1}" = B ] && bin=$B
  taskset -c 12 "$bin" --case xlsx_one_cell_commit_save,xlsx_one_percent_commit_save,xlsx_noop_commit_save \
    --xlsx-shape dense-wide,medium --samples 5 --warmup 2 --json "$OUT/alloc-$slot.json" > /dev/null 2> "$OUT/alloc-$slot.stderr"
  echo "$(date +%T) alloc $slot done"
done
