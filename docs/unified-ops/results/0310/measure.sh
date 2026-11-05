#!/bin/bash
# Record 0310 measurement driver.
#
# BASE holds binaries built from trunk 0f3d8158cb, AFTER from this record's
# working tree (SHA-256 in binaries.sha256):
#   unified-query[-all]   cargo build --release --manifest-path
#                         tools/perf-baseline/unified-query/Cargo.toml [--features all-formats]
#   unified-rtf-*         cargo build --release --manifest-path
#                         tools/perf-baseline/<harness>/Cargo.toml (root Cargo.lock copied
#                         next to the manifest, not committed)
#   <example>             cargo build --profile profiling -p litchi-ops --features
#                         docx,pptx,xlsx,xlsb,xls,doc,ppt,rtf,odt,ods,odp,odg,odc,odi,odb,oth,odm,odf-formula
#                         --example <example>   (stripped before timing)
# `trunk` runs every harness twice on the base binaries (the current-trunk
# re-measurement); `abba` runs the harnesses whose routes this record changes
# (unified-query, xls_p0_write_bench) in B A A B process order. Each harness is itself an
# in-process ABBA. One pinned CPU; perf is blocked (perf_event_paranoid=4),
# instruction counts are callgrind.
set -u
MODE=$1 BASE=$2 AFTER=${3:-} OUT=$4 CPU=${CPU:-31}
mkdir -p "$OUT"
export UNIFIED_QUERY_DIR=${UNIFIED_QUERY_DIR:-$OUT/scratch}; mkdir -p "$UNIFIED_QUERY_DIR"
log() { echo "== $* $(date +%T) load $(cut -d' ' -f1 /proc/loadavg)" >> "$OUT/log.txt"; }
run() { local name=$1; shift; log "$name"; taskset -c "$CPU" "$@" > "$OUT/$name" 2>> "$OUT/log.txt" || echo "FAIL $name" >> "$OUT/log.txt"; }
FAMILIES="xlsx_1000_bytes xlsx_1000_file xlsx_50000_bytes xlsx_50000_file xlsb_1000_bytes xlsb_1000_file pptx_small_bytes pptx_small_file pptx_large_bytes pptx_large_file docx_drawing_file docx_saut_bytes docx_saut_file"
case $MODE in
trunk)
  for i in 1 2; do
    run rtf_destination-$i.json "$BASE/rtf_destination_bench"
    run xls_read-$i.json "$BASE/xls_read_bench"
    run ods-$i.json env LITCHI_ODS_BENCH_SAMPLES=9 "$BASE/ods_unified_bench"
    run odt-$i.txt "$BASE/odt_paired_bench"
    run docx_fanout-$i.json env DOCX_FANOUT_SIZES=1,100 "$BASE/docx_fanout_bench"
    run docx_element-$i.json "$BASE/docx_element_set_bench"
    run checklist-$i.jsonl env CHECKLIST_SCENARIO=odp_metadata_read "$BASE/checklist_fixes_bench"
    run unified-rtf-add-$i.json "$BASE/unified-rtf-add"
    run unified-rtf-run-add-$i.json "$BASE/unified-rtf-run-add"
    run unified-rtf-layout-$i.json "$BASE/unified-rtf-layout"
    run unified-rtf-batch-set-$i.json "$BASE/unified-rtf-batch-set"
    run xls_p0-$i.json "$BASE/xls_p0_write_bench"
  done
  for i in 1 2; do
    for fam in $FAMILIES; do
      log "query-$i $fam"
      taskset -c "$CPU" "$BASE/unified-query" time "$fam" 9 >> "$OUT/query-$i.jsonl" 2>> "$OUT/log.txt" || echo "FAIL query $fam" >> "$OUT/log.txt"
    done
    for fam in xlsx_1000_bytes pptx_large_bytes; do
      log "query-all-$i $fam"
      taskset -c "$CPU" "$BASE/unified-query-all" time "$fam" 9 >> "$OUT/query-all-$i.jsonl" 2>> "$OUT/log.txt" || echo "FAIL query-all $fam" >> "$OUT/log.txt"
    done
  done
  # iWork (features pages,keynote,numbers; the same profiling build).
  for i in 1 2; do
    run iwork-directory-$i.jsonl "$BASE/iwork_directory_bench" --rounds 8 --iterations 30
    run keynote-$i.jsonl "$BASE/keynote_bench" --rounds 8 --iterations 30
    run numbers-$i.json env LITCHI_NUMBERS_BENCH_ONLY=open_get LITCHI_NUMBERS_BENCH_SAMPLES=7 "$BASE/numbers_unified_bench"
  done
  for i in 1 2 3; do log text_set-$i.json; taskset -c "$CPU" "$BASE/text_set_bench" "$OUT/text_set-$i.json" > /dev/null; done
  ;;
abba)
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
    run xls_p0-$leg.json "$dir/xls_p0_write_bench"
  done
  ;;
esac
echo "done $MODE" >> "$OUT/log.txt"
