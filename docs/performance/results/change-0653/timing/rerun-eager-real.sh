#!/usr/bin/env bash
# change 0653: the eager-real scenario alone, re-measured in more, shorter
# interleaved blocks after the first window showed a 317% A/A spread from host
# contention (load average 25-45 on 32 cores, seven other measuring agents).
set -u
S="$1"; B="$2"; A="$3"; CPU="${4:-8}"; N="${5:-20}"
OUT="$S/timing"; mkdir -p "$OUT"
run() { taskset -c "$CPU" "$1" time eager "$S/real.xlsx" H680 3 "$N"; }
rm -f "$OUT"/eager-real.*.txt
for r in 1 2 3 4 5 6; do
  run "$B" > "$OUT/eager-real.before.$r.txt"
  run "$A" > "$OUT/eager-real.after.$((r*2-1)).txt"
  run "$A" > "$OUT/eager-real.after.$((r*2)).txt"
  run "$B" > "$OUT/eager-real.floorA.$r.txt"
done
echo "eager-real rerun done"
