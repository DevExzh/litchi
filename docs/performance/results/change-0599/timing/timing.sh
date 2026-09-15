#!/usr/bin/env bash
# Paired timing for change 0599, order A1 B1 B2 A2 plus an A/A control pair
# A3 A4 taken in the same window. Every leg is pinned to one CPU.
set -euo pipefail
SC="$1"; REPO="$2"; CPU="$3"
BEFORE=/home/zhuhe/code/litchi-worktrees/targets/0599-before/release/xlsb_crud
AFTER=/home/zhuhe/code/litchi-worktrees/0599/tools/perf-baseline/target/release/xlsb_crud
WARMUP=3
SAMPLES=40
cd "$REPO"
declare -A FIX=(
  [poi]="$REPO/test-data/poi/test-data/spreadsheet/testVarious.xlsb"
  [cond]="$REPO/test-data/ooxml/xlsb/cond_format.xlsb"
  [syn_small]="$SC/synthetic-4x500x8.xlsb"
  [syn_large]="$SC/synthetic-4x2000x12.xlsb"
)
for leg in A1 B1 B2 A2 A3 A4; do
  case "$leg" in A*) BIN="$BEFORE";; B*) BIN="$AFTER";; esac
  for tag in poi cond syn_small syn_large; do
    taskset -c "$CPU" "$BIN" --case all --warmup "$WARMUP" --samples "$SAMPLES" \
      --fixture "${FIX[$tag]}" --json "$SC/timing/${leg}-${tag}.json" > /dev/null
  done
  echo "leg $leg done"
done
