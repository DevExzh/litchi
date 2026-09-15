#!/usr/bin/env bash
# Paired ABBA timing for change 0611, plus an A/A floor in the same window.
#
# Usage: run-timing.sh <work-dir>
#
# Leg order is A1 B1 B2 A2 for the paired comparison and A3 A4 for the floor,
# all pinned to CPU 31, both binaries built --release --locked from the same
# sources with the same flags.
set -euo pipefail
WORK="${1:?usage: run-timing.sh <work-dir>}"
BEFORE_BIN=/home/zhuhe/code/litchi-worktrees/targets/0611-before/release/litchi-perf-baseline
AFTER_BIN=/home/zhuhe/code/litchi-worktrees/targets/0611-after/release/litchi-perf-baseline
BEFORE_CWD=/home/zhuhe/code/litchi-worktrees/before-2d6fbeaed
AFTER_CWD=/home/zhuhe/code/litchi-worktrees/0611
CPU=31
SAMPLES=30
WARMUP=5

RANGE_CASES=opc_range_source_open,opc_range_source_open_main_read,xlsx_range_source_open,xlsx_range_source_first_cell
# Change 0493's and 0572's transport: 1 ms of service per request, 100 MiB/s,
# 64 KiB maximum physical range.
RANGE_FLAGS=(--range-fixed-latency-us 1000 --range-request-overhead-us 0
             --range-bandwidth-bytes-per-sec 104857600 --range-max-physical-bytes 65536)

LOCAL_CASES=opc_file_source_open,docx_file_source_open,pptx_file_source_open,docx_file_source_full_text

mkdir -p "$WORK/fsroot"

run_leg() { # <bin> <cwd> <label> <kind>
  local bin="$1" cwd="$2" label="$3" kind="$4"
  case "$kind" in
    range)
      ( cd "$cwd" && taskset -c "$CPU" "$bin" \
          --warmup "$WARMUP" --samples "$SAMPLES" --shape tiny --payload compressible \
          --xlsx-shape tiny --case "$RANGE_CASES" "${RANGE_FLAGS[@]}" \
          --json "$WORK/range-$label.json" >/dev/null )
      ;;
    local)
      ( cd "$cwd" && taskset -c "$CPU" "$bin" \
          --warmup "$WARMUP" --samples "$SAMPLES" --shape tiny --payload compressible \
          --filesystem-cache warm --filesystem-root "$WORK/fsroot" \
          --case "$LOCAL_CASES" --json "$WORK/local-$label.json" >/dev/null )
      ;;
  esac
  echo "  leg $label ($kind) done"
}

for kind in range local; do
  echo "== $kind: A1 B1 B2 A2, then the A/A floor A3 A4"
  run_leg "$BEFORE_BIN" "$BEFORE_CWD" "${kind}-A1" "$kind"
  run_leg "$AFTER_BIN"  "$AFTER_CWD"  "${kind}-B1" "$kind"
  run_leg "$AFTER_BIN"  "$AFTER_CWD"  "${kind}-B2" "$kind"
  run_leg "$BEFORE_BIN" "$BEFORE_CWD" "${kind}-A2" "$kind"
  run_leg "$BEFORE_BIN" "$BEFORE_CWD" "${kind}-A3" "$kind"
  run_leg "$BEFORE_BIN" "$BEFORE_CWD" "${kind}-A4" "$kind"
done
echo "done: $WORK"
