#!/bin/bash
# Record 0315 measurement driver.
#
# BASE holds binaries built from trunk 30eed270b5 with this record's new
# routes only (harness and example sources at this record, library at the
# base); AFTER from this record's tree. SHA-256 in binaries.sha256.
#   unified-query[-all]  cargo build --release --offline --manifest-path
#                        tools/perf-baseline/unified-query/Cargo.toml [--features all-formats]
#   <example>            cargo build --profile profiling -p litchi-ops --no-default-features
#                        --features F --example <example>
#                        (F: ods, rtf, odt; xls_read_bench with every non-iWork format)
# Processes run B A A B on one pinned CPU; each harness is itself an
# in-process ABBA. perf is blocked (perf_event_paranoid=4); instruction
# counts are callgrind (callgrind-summary.txt).
set -u
BASE=$1 AFTER=$2 OUT=$3 CPU=${CPU:-31}
mkdir -p "$OUT"
export UNIFIED_QUERY_DIR=${UNIFIED_QUERY_DIR:-$OUT/scratch}; mkdir -p "$UNIFIED_QUERY_DIR"
log() { echo "== $* $(date +%T) load $(cut -d' ' -f1 /proc/loadavg)" >> "$OUT/log.txt"; }
run() { local name=$1; shift; log "$name"; taskset -c "$CPU" "$@" > "$OUT/$name" 2>> "$OUT/log.txt" || echo "FAIL $name" >> "$OUT/log.txt"; }
FAMILIES="xlsx_1000_bytes xlsx_1000_file xlsx_50000_bytes xlsx_50000_file xlsb_1000_bytes xlsb_1000_file pptx_small_bytes pptx_small_file pptx_large_bytes pptx_large_file docx_drawing_bytes docx_drawing_file docx_saut_bytes docx_saut_file"
for leg in base-1 after-1 after-2 base-2; do
  dir=$BASE; case $leg in after*) dir=$AFTER;; esac
  for fam in $FAMILIES; do
    log "query-$leg $fam"
    taskset -c "$CPU" "$dir/unified-query" time "$fam" 9 >> "$OUT/query-$leg.jsonl" 2>> "$OUT/log.txt" || echo "FAIL query $fam" >> "$OUT/log.txt"
  done
  for fam in xlsx_1000_bytes pptx_large_bytes; do
    log "query-all-$leg $fam"
    taskset -c "$CPU" "$dir/unified-query-all" time "$fam" 9 >> "$OUT/query-all-$leg.jsonl" 2>> "$OUT/log.txt" || echo "FAIL query-all $fam" >> "$OUT/log.txt"
  done
  run xls_read-$leg.json "$dir/xls_read_bench"
  run ods-$leg.json env LITCHI_ODS_BENCH_SAMPLES=9 LITCHI_ODS_BENCH_ONLY=query,query_warm,query_cold,query_warm_owned "$dir/ods_unified_bench"
  run odt-$leg.txt "$dir/odt_paired_bench"
  run rtf_destination-$leg.json "$dir/rtf_destination_bench"
  run unified-rtf-layout-$leg.json "$dir/rtf_layout_bench"
  run unified-rtf-batch-set-$leg.json env RTF_BATCH_SET_COUNT=1 "$dir/rtf_batch_set_bench"
done
rm -rf "$UNIFIED_QUERY_DIR"
echo "done" >> "$OUT/log.txt"
