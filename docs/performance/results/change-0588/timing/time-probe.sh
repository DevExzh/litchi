#!/usr/bin/env bash
# change 0588 paired timing: ABBA over the public XLSX reads on the real
# producer fixture and its marker-stripped control, plus an A/A floor block
# measured in the same window. One process per block; CPU-pinned.
set -u
S="$1"; B="$2"; A="$3"; CPU="${4:-8}"; N="${5:-40}"
OUT="$S/timing"; mkdir -p "$OUT"
run() { taskset -c "$CPU" "$1" time "$2" "$3" H680 5 "$N"; }
for scenario in "eager $S/real.xlsx" "source $S/real.xlsx" "eager $S/control.xlsx"; do
  set -- $scenario
  mode="$1"; file="$2"
  tag="$mode-$(basename "$file" .xlsx)"
  run "$B" "$mode" "$file" > "$OUT/$tag.before.1.txt"
  run "$A" "$mode" "$file" > "$OUT/$tag.after.1.txt"
  run "$A" "$mode" "$file" > "$OUT/$tag.after.2.txt"
  run "$B" "$mode" "$file" > "$OUT/$tag.before.2.txt"
  # A/A floor in the same window
  run "$B" "$mode" "$file" > "$OUT/$tag.floorA.1.txt"
  run "$B" "$mode" "$file" > "$OUT/$tag.floorB.1.txt"
  run "$B" "$mode" "$file" > "$OUT/$tag.floorB.2.txt"
  run "$B" "$mode" "$file" > "$OUT/$tag.floorA.2.txt"
done
echo "timing done"
