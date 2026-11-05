#!/bin/sh
# Record 0305 measurement driver. BASE and AFTER hold the `profiling`
# examples built from f24868dfb5 (trunk; the merged trunk e7fe30fb24 changed
# no source on these routes) and from the record's merged branch with
#   cargo build --profile profiling -p litchi-ops \
#     --features docx,pages,numbers,keynote,rtf --example <name>
# (debug info stripped; SHA-256 in binaries.sha256). Processes run B A A B
# on one pinned CPU; each harness is itself an in-process ABBA.
set -eu
BASE=$1 AFTER=$2 OUT=$3 CPU=${CPU:-30}
mkdir -p "$OUT"
run() { taskset -c "$CPU" "$@"; }
for leg in base-1 after-1 after-2 base-2; do
  dir=$BASE; case $leg in after*) dir=$AFTER;; esac
  run "$dir/keynote_bench" --rounds 8 --iterations 30 | grep summary > "$OUT/keynote-$leg.jsonl"
  LITCHI_NUMBERS_BENCH_SAMPLES=7 run "$dir/numbers_unified_bench" > "$OUT/numbers-$leg.json"
  DOCX_FANOUT_SIZES=1,100 run "$dir/docx_fanout_bench" > "$OUT/docx-$leg.json"
  for s in batch_1 batch_1000 batch_10000; do
    run "$dir/pages_bench" --only $s --rounds 4 --iterations 12 | tail -1
  done > "$OUT/pages-$leg.jsonl"
  run "$dir/rtf_destination_bench" > "$OUT/rtf-$leg.json"
done
