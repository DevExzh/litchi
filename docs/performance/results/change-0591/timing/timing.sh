#!/bin/bash
# Paired timing for change 0591, order A1 B1 B2 A2 in one window, pinned to CPU 11.
set -u
BEFORE=/home/zhuhe/code/litchi-worktrees/targets/0591-before/release/litchi-perf-baseline
AFTER=/home/zhuhe/code/litchi-worktrees/targets/0591-after/release/litchi-perf-baseline
OUT="$1"
CASES=docx_semantic_one_edit_save,docx_semantic_one_percent_edit_save,docx_semantic_noop_edit_save
mkdir -p "$OUT"
run() {
  rm -f "$OUT/$2.json"
  taskset -c 11 "$1" --warmup 5 --samples 50 --case "$CASES" --json "$OUT/$2.json" > "$OUT/$2.log" 2>&1
  echo "$2 exit=$?"
}
run "$BEFORE" A1
run "$AFTER"  B1
run "$AFTER"  B2
run "$BEFORE" A2
