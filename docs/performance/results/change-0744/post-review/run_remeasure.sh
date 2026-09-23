#!/usr/bin/env bash
# Change 0744 post-review re-measurement: dense-wide ABBA with both legs
# rebuilt by the identical command (before: 009d515bef from the equal-length
# worktree; after: a755ff8c26), CPU 12, order A1 B1 B2 A2 A3 B3 B4 A4.
set -euo pipefail
A=${A:-/home/zhuhe/code/litchi-worktrees/targets/0744-before/release/litchi-perf-baseline}
B=${B:-/home/zhuhe/code/litchi-worktrees/targets/0744/release/litchi-perf-baseline}
OUT=${OUT:-raw}
mkdir -p "$OUT"
for slot in A1 B1 B2 A2 A3 B3 B4 A4; do
  bin=$A; [ "${slot:0:1}" = B ] && bin=$B
  taskset -c 12 "$bin" --case xlsx_first_cell,xlsx_full_cell_scan,xlsx_one_cell_commit_save,xlsx_one_percent_commit_save \
    --xlsx-shape dense-wide --samples 20 --warmup 3 --json "$OUT/dense-$slot.json" > /dev/null 2> "$OUT/dense-$slot.stderr"
  echo "$(date +%T) dense $slot done"
done
