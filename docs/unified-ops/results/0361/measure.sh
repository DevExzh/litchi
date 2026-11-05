#!/bin/bash
# Record 0361 measurements. Base = trunk f8899dd2cb, after = uo2/w7-iwork
# (the trunk merge). One pinned CPU; processes in the order B A A B.
#  - keynote_bench / pages_bench / numbers_unified_bench (record 0329's
#    in-process paired leaf/unified benches: open, set, structural edits,
#    1k batch; numbers includes record 0347's sheet/table swap scenarios);
#  - unified-batch-closure (record 0337/0340 harness): pages, keynote,
#    numbers 1k- and 10k-op batches, 5-round in-process ABBA, fixture of 10,000.
# Usage: measure.sh BIN_DIR OUT_DIR (binaries <bench>.base / <bench>.after).
set -eu
BIN=$1 OUT=$2 CPU=${CPU:-30}
mkdir -p "$OUT/abba"
run() {
  label=$1; v=$2
  echo "== $label $(date +%T) load $(cut -d' ' -f1 /proc/loadavg)" >> "$OUT/log.txt"
  taskset -c "$CPU" "$BIN/keynote_bench.$v" --rounds 8 --iterations 30 > "$OUT/keynote-$label.jsonl"
  taskset -c "$CPU" "$BIN/pages_bench.$v" --rounds 8 --iterations 30 > "$OUT/pages-$label.jsonl"
  LITCHI_NUMBERS_BENCH_SAMPLES=8 taskset -c "$CPU" "$BIN/numbers_unified_bench.$v" > "$OUT/numbers-$label.json"
  taskset -c "$CPU" "$BIN/batch.$v" --formats pages,keynote,numbers --counts 1000,10000 --rounds 5 > "$OUT/abba/batch-A-$v-${label%%-*}.json"
}
run 1-base base
run 2-after after
run 3-after after
run 4-base base
sha256sum "$BIN"/*.base "$BIN"/*.after | sed "s|$BIN/||" > "$OUT/binaries.sha256"
echo done >> "$OUT/log.txt"
