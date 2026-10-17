#!/bin/bash
# Record 0326 measurement driver. BASE = binaries built from trunk
# 16b800350b (plus the new path_parse_bench example source), AFTER = the
# merged branch tree; build commands in build.sh, SHA-256 in binaries.sha256.
# Processes run B A A B on one pinned CPU; every harness is itself an
# in-process ABBA. perf is blocked (perf_event_paranoid=4).
set -u
BASE=$1 AFTER=$2 OUT=$3 CPU=${CPU:-30}
mkdir -p "$OUT"
export UNIFIED_QUERY_DIR=$OUT/scratch; mkdir -p "$UNIFIED_QUERY_DIR"
log() { echo "== $* $(date +%T) load $(cut -d' ' -f1 /proc/loadavg)" >> "$OUT/log.txt"; }
run() { local name=$1; shift; log "$name"; taskset -c "$CPU" "$@" >> "$OUT/$name" 2>> "$OUT/log.txt" || echo "FAIL $name" >> "$OUT/log.txt"; }
for leg in base-1 after-1 after-2 base-2; do
  dir=$BASE; case $leg in after*) dir=$AFTER;; esac
  for fam in pptx_small_bytes pptx_small_file pptx_large_bytes pptx_large_file; do
    run query-$leg.jsonl "$dir/unified-query" time "$fam" 9
  done
  run path-$leg.jsonl "$dir/path_parse_bench" --rounds 9 --iterations 200000
  for s in batch_1000 batch_1000_parsed batch_10000 batch_10000_parsed; do
    run pages-$leg.jsonl "$dir/pages_bench" --rounds 6 --iterations 30 --only $s
  done
  for s in batch_1000 batch_1000_parsed; do
    run keynote-$leg.jsonl "$dir/keynote_bench" --rounds 6 --iterations 30 --only $s
  done
  run numbers-$leg.json env LITCHI_NUMBERS_BENCH_SAMPLES=8 LITCHI_NUMBERS_BENCH_ONLY=batch_1000,batch_1000_owned,batch_1000_parsed "$dir/numbers_unified_bench"
  run unified-rtf-batch-set-$leg.json env RTF_BATCH_SET_COUNT=1 "$dir/rtf_batch_set_bench"
  run unified-rtf-layout-$leg.json "$dir/rtf_layout_bench"
  for c in style_get_format paragraph_list_format; do
    run docx-$c-$leg.json env DOCX_BENCH_CASE=$c "$dir/docx_parts_bench"
  done
  run xlsx-$leg.jsonl env FEATURE_SCENARIO=comment_get "$dir/xlsx_features_bench"
  run xlsb-$leg.jsonl env FEATURE_SCENARIO=comment_get "$dir/xlsb_features_bench"
  run pptx-elements-$leg.json "$dir/unified-pptx-elements" --rounds 9
done
rm -rf "$UNIFIED_QUERY_DIR"
echo "done" >> "$OUT/log.txt"
