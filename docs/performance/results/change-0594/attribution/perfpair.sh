#!/usr/bin/env bash
# perf stat isolation pair: run the probe with N and 2N opens, difference the
# counters and divide by N to get the per-open cost. Events are split into two
# groups so nothing multiplexes.
set -euo pipefail
S="$1"; BIN="$2"; FIX="$3"; LEG="$4"; N="$5"; EVENTS="$6"; GROUP="$7"
for R in "$N" "$((N*2))"; do
  taskset -c 14 perf stat -e "$EVENTS" -x, -r 5 "$BIN" "$FIX" "$R" \
    > /dev/null 2> "$S/perf-$LEG-$GROUP-$R.csv"
done
