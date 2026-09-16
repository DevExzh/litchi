#!/usr/bin/env bash
# Paired ABBA timing for change 0632, plus an A/A floor in the same window.
#
# Usage: run-timing.sh <work-dir>
#
# Leg order is A1 B1 B2 A2 for the paired comparison and A3 A4 for the floor,
# all pinned to CPU 8, both binaries built --release --locked from the same
# sources with the same flags.  Change 0627 found a concurrent build relinking
# a binary mid-run, so both are copied out of their Cargo target directories
# before the first leg and every leg runs the copy.
set -euo pipefail
WORK="${1:?usage: run-timing.sh <work-dir>}"
BEFORE_TARGET=/home/zhuhe/code/litchi-worktrees/targets/0632-before/release/litchi-perf-baseline
AFTER_TARGET=/home/zhuhe/code/litchi-worktrees/targets/0632-after/release/litchi-perf-baseline
BEFORE_CWD=/home/zhuhe/code/litchi-worktrees/before-c7326f680
AFTER_CWD=/home/zhuhe/code/litchi-worktrees/0632
CPU=8
SAMPLES=30
WARMUP=5

mkdir -p "$WORK/bin" "$WORK/fsroot"
cp "$BEFORE_TARGET" "$WORK/bin/litchi-perf-baseline-before"
cp "$AFTER_TARGET"  "$WORK/bin/litchi-perf-baseline-after"
BEFORE_BIN="$WORK/bin/litchi-perf-baseline-before"
AFTER_BIN="$WORK/bin/litchi-perf-baseline-after"
sha256sum "$BEFORE_BIN" "$AFTER_BIN" | tee "$WORK/binaries.sha256"

# The ZIP-backed range-source selectors, plus one OLE2 selector as a control:
# `xls_range_source_open` does not go through the ZIP locator at all.
RANGE_CASES=opc_range_source_open,opc_range_source_open_main_read,xlsx_range_source_open,xlsx_range_source_first_cell,xlsx_range_source_list_sheets,xls_range_source_open
OLE2_FILE=test-data/ole/xls/WithCustomViews.xls
# Change 0493's and 0572's transport: 1 ms of service per request, 100 MiB/s,
# 64 KiB maximum physical range.
RANGE_FLAGS=(--range-fixed-latency-us 1000 --range-request-overhead-us 0
             --range-bandwidth-bytes-per-sec 104857600 --range-max-physical-bytes 65536)

LOCAL_CASES=opc_file_source_open,docx_file_source_open,pptx_file_source_open,docx_file_source_full_text

run_leg() { # <bin> <cwd> <label> <kind>
  local bin="$1" cwd="$2" label="$3" kind="$4"
  case "$kind" in
    range)
      ( cd "$cwd" && taskset -c "$CPU" "$bin" \
          --warmup "$WARMUP" --samples "$SAMPLES" --shape tiny --payload compressible \
          --xlsx-shape tiny --case "$RANGE_CASES" --ole2-file "$OLE2_FILE" \
          "${RANGE_FLAGS[@]}" \
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
