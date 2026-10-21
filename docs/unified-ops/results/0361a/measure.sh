#!/bin/bash
# Record 0361a: open routes (and the rest of each paired bench) of base
# 7138746bb7 vs after, pinned CPU, processes B A A B B A A B.
set -eu
BIN=$1 OUT=$2 CPU=${CPU:-30}
mkdir -p "$OUT"
run() {
  label=$1; v=$2
  echo "== $label $(date +%T) load $(cut -d' ' -f1 /proc/loadavg)" >> "$OUT/log.txt"
  taskset -c "$CPU" "$BIN/keynote_bench.$v" --rounds 8 --iterations 30 > "$OUT/keynote-$label.jsonl"
  taskset -c "$CPU" "$BIN/pages_bench.$v" --rounds 8 --iterations 30 > "$OUT/pages-$label.jsonl"
  LITCHI_NUMBERS_BENCH_SAMPLES=8 taskset -c "$CPU" "$BIN/numbers_unified_bench.$v" > "$OUT/numbers-$label.json"
}
run 1-base base; run 2-after after; run 3-after after; run 4-base base
run 5-base base; run 6-after after; run 7-after after; run 8-base base
sha256sum "$BIN"/*.base "$BIN"/*.after | sed "s|$BIN/||" > "$OUT/binaries.sha256"
echo done >> "$OUT/log.txt"
