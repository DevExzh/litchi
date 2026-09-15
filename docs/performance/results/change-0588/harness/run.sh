#!/usr/bin/env bash
# change 0588: ABBA over four existing XLSX selectors. The harness corpora carry
# no MCE markers (0587 section 4), so these are the no-regression leg for the
# codec's borrowed fast path, not a win.
set -u
S="$1"; HB="$2"; HA="$3"; CASES="$4"; CPU="${5:-8}"
O="$S/harness"; mkdir -p "$O"
run() { taskset -c "$CPU" "$1" --case "$CASES" --warmup 3 --samples 30 --json "$O/$2.json" > /dev/null 2>&1; }
run "$HB" before-1
run "$HA" after-1
run "$HA" after-2
run "$HB" before-2
run "$HB" floorA
run "$HB" floorB
echo harness done
