#!/bin/bash
# Record 0287 paired measurements: processes in B A A B order, each an
# in-process ABBA (leaf vs unified). Pinned to CPU 31.
set -u
S=/deleted-coord-artifact
B=$S/bin
R=$S/results
export UNIFIED_QUERY_DIR=$S/uq
run() { # name cmd...
  local name=$1; shift
  echo "== $name $(date +%T) load $(cut -d' ' -f1 /proc/loadavg)" >> $R/log.txt
  "$@" > $R/$name 2>> $R/log.txt
}
i=1
for side in base after after base; do
  for fam in xlsx_1000_bytes xlsx_1000_file pptx_small_bytes pptx_large_bytes pptx_large_file docx_drawing_file xlsb_1000_file xls_1000; do
    taskset -c 31 $B/unified-query.$side time $fam 9 >> $R/query-$i-$side.jsonl 2>> $R/log.txt
  done
  echo "== query-$i-$side done $(date +%T)" >> $R/log.txt
  i=$((i+1))
done
i=1
for side in base after after base; do
  run query-all-$i-$side.jsonl taskset -c 31 $B/unified-query-all.$side time xlsx_1000_bytes 9
  run query-all-pptx-$i-$side.jsonl taskset -c 31 $B/unified-query-all.$side time pptx_large_bytes 9
  i=$((i+1))
done
i=1
for side in base after after base; do
  run ods-$i-$side.json env LITCHI_ODS_BENCH_SAMPLES=9 taskset -c 31 $B/ods_unified_bench.$side
  run odt-$i-$side.txt taskset -c 31 $B/odt_paired_bench.$side
  run rtf-$i-$side.json taskset -c 31 $B/rtf_destination_bench.$side
  run xls-$i-$side.json env LITCHI_XLS_BENCH_OWNER=positional taskset -c 31 $B/xls_read_bench.$side
  i=$((i+1))
done
echo "done $(date +%T)" >> $R/log.txt
