#!/usr/bin/env bash
# change 0653: the eager-real scenario, re-measured in a quieter window after
# the first two windows showed A/A spreads of 317% and 623% from host
# contention. Six ABBA rounds, 20 samples each, with one before-leg floor block
# per round.
set -u
S="$1"; B="$2"; A="$3"; CPU="${4:-8}"; N="${5:-20}"
OUT="$S/timing2"; mkdir -p "$OUT"
run() { taskset -c "$CPU" "$1" time eager "$S/real.xlsx" H680 3 "$N"; }
for r in 1 2 3 4 5 6; do
  run "$B" > "$OUT/eager-real.before.$r.txt"
  run "$A" > "$OUT/eager-real.after.$((r*2-1)).txt"
  run "$A" > "$OUT/eager-real.after.$((r*2)).txt"
  run "$B" > "$OUT/eager-real.floorA.$r.txt"
done
echo "eager-real window 3 done"
