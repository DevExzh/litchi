#!/bin/sh
# Record 0302 measurement driver. BASE and AFTER hold the `profiling`
# examples built from 928b19b4db and from 7f41bba583 (m4_formats_bench from
# b046c2a64c; its ODT rows were re-run B A A B after that commit) with
#   cargo build --profile profiling -p litchi-ops \
#     --features docx,odt,doc,ods,pages,numbers,rtf --example <name>
# (debug info stripped afterwards; SHA-256 in binaries.sha256). Processes run
# B A A B on one pinned CPU; each harness is itself an in-process ABBA.
set -eu
BASE=$1 AFTER=$2 OUT=$3 CPU=${CPU:-30}
mkdir -p "$OUT"
run() { taskset -c "$CPU" "$@"; }
for leg in base-1 after-1 after-2 base-2; do
  dir=$BASE; case $leg in after*) dir=$AFTER;; esac
  DOCX_FANOUT_SIZES=1,100 run "$dir/docx_fanout_bench" > "$OUT/docx-$leg.json"
  M4_ONLY='odt*' M4_SIZE=10000 run "$dir/m4_formats_bench" > "$OUT/odt-$leg.txt"
  for s in batch_1 batch_1000 batch_10000; do
    run "$dir/pages_bench" --only $s --rounds 4 --iterations 12 | tail -1
  done > "$OUT/pages-$leg.jsonl"
  LITCHI_NUMBERS_BENCH_ONLY=open_get,set,batch_1000 LITCHI_NUMBERS_BENCH_SAMPLES=7 run "$dir/numbers_unified_bench" > "$OUT/numbers-$leg.json"
  run "$dir/rtf_destination_bench" > "$OUT/rtf-$leg.json"
done
