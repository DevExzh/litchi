#!/usr/bin/env bash
# ABBA paired timing for change 0600. Order A1 B1 B2 A2; identical flags,
# identical CPU (22), one leg at a time.
set -euo pipefail
S="$1"; WARMUP="$2"; SAMPLES="$3"; TAG="$4"
BEFORE_BIN=/home/zhuhe/code/litchi-worktrees/targets/0600-before/release/litchi-perf-baseline
AFTER_BIN=/home/zhuhe/code/litchi-worktrees/0600/tools/perf-baseline/target/release/litchi-perf-baseline
BEFORE_WD=/home/zhuhe/code/litchi-worktrees/before-8fe9efa55
AFTER_WD=/home/zhuhe/code/litchi-worktrees/0600
CASES=docx_file_source_open,pptx_file_source_open
run() {
  local leg="$1" bin="$2" wd="$3"
  ( cd "$wd" && taskset -c 22 "$bin" --warmup "$WARMUP" --samples "$SAMPLES" \
      --case "$CASES" --json "$S/abba-$TAG-$leg.json" ) > "$S/abba-$TAG-$leg.log" 2>&1
  echo "$leg done"
}
run A1 "$BEFORE_BIN" "$BEFORE_WD"
run B1 "$AFTER_BIN"  "$AFTER_WD"
run B2 "$AFTER_BIN"  "$AFTER_WD"
run A2 "$BEFORE_BIN" "$BEFORE_WD"
