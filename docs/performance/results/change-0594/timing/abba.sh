#!/usr/bin/env bash
# ABBA paired timing for change 0594. Order A1 B1 B2 A2; identical flags,
# identical CPU, one leg at a time.
set -euo pipefail
S="$1"; WARMUP="$2"; SAMPLES="$3"; TAG="$4"
BEFORE_BIN=/home/zhuhe/code/litchi-worktrees/targets/0594-before/release/litchi-perf-baseline
AFTER_BIN=/home/zhuhe/code/litchi-worktrees/targets/0594-after/release/litchi-perf-baseline
BEFORE_WD=/home/zhuhe/code/litchi-worktrees/before-08d968f8e
AFTER_WD=/home/zhuhe/code/litchi-worktrees/0594
CASES=docx_file_source_open,pptx_file_source_open,xlsx_source_open
run() { # leg bin wd
  local leg="$1" bin="$2" wd="$3"
  ( cd "$wd" && taskset -c 14 "$bin" --warmup "$WARMUP" --samples "$SAMPLES" \
      --case "$CASES" --xlsx-shape medium --json "$S/abba-$TAG-$leg.json" ) \
      > "$S/abba-$TAG-$leg.log" 2>&1
  echo "$leg done"
}
run A1 "$BEFORE_BIN" "$BEFORE_WD"
run B1 "$AFTER_BIN"  "$AFTER_WD"
run B2 "$AFTER_BIN"  "$AFTER_WD"
run A2 "$BEFORE_BIN" "$BEFORE_WD"
