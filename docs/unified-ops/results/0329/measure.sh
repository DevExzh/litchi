#!/bin/sh
# Record 0329 measurements: base (fc274041da) and after builds of the three
# iWork benches, each with only its format feature
# (`cargo build --release -p litchi-ops --features <format> --example <bench>`),
# pinned to one core, processes in the order B A A B.
# Usage: measure.sh BIN_DIR OUT_DIR (binaries named <bench>.base / <bench>.after).
set -eu
BIN="$1"
OUT="$2"
mkdir -p "$OUT"
run() {
    label="$1"; version="$2"
    taskset -c 7 "$BIN/keynote_bench.$version" --rounds 8 --iterations 30 > "$OUT/keynote-$label.jsonl"
    taskset -c 7 "$BIN/pages_bench.$version" --rounds 8 --iterations 30 > "$OUT/pages-$label.jsonl"
    LITCHI_NUMBERS_BENCH_SAMPLES=8 taskset -c 7 "$BIN/numbers_unified_bench.$version" > "$OUT/numbers-$label.json"
}
run 1-base base
run 2-after after
run 3-after after
run 4-base base
sha256sum "$BIN"/*.base "$BIN"/*.after | sed "s|$BIN/||" > "$OUT/binaries.sha256"
