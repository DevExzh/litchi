#!/bin/sh
# Record 0291 measurement driver. BASE and AFTER are directories holding the
# `profiling`-profile examples built from 99cc3138eb and from this record's
# commits with
#   cargo build --profile profiling -p litchi-ops \
#     --features docx,pptx,xlsx,xlsb,xls,doc,ppt,rtf,odt,ods,odp,odg,odc,odi,odb,oth,odm,odf-formula \
#     --example <name>
# Processes run in B A A B order on one pinned CPU; each harness is itself an
# in-process ABBA. Callgrind counts are deterministic (perf is unavailable).
set -eu
BASE=$1 AFTER=$2 OUT=$3 CPU=${CPU:-31}
mkdir -p "$OUT"
run() { taskset -c "$CPU" "$@"; }
for bench in rtf_destination_bench odt_paired_bench xls_read_bench ods_unified_bench docx_fanout_bench odf_gaps_bench; do
  for leg in base-1 after-1 after-2 base-2; do
    dir=$BASE; case $leg in after*) dir=$AFTER;; esac
    case $bench in
      docx_fanout_bench) DOCX_FANOUT_SIZES=1,100 run "$dir/$bench" > "$OUT/$bench-$leg.json" ;;
      *) run "$dir/$bench" > "$OUT/$bench-$leg.json" ;;
    esac
  done
done
# Instruction counts.
for tag in base after; do
  dir=$BASE; [ $tag = after ] && dir=$AFTER
  for n in 1000 3000; do for leg in unified leaf; do
    DOCX_FANOUT_SIZES=1 DOCX_FANOUT_PROFILE=$leg DOCX_FANOUT_LOOPS=$n run valgrind --tool=callgrind \
      --callgrind-out-file="$OUT/cg-docx-$tag-$leg-$n.out" "$dir/docx_fanout_bench" > /dev/null
  done; done
  LITCHI_ODS_BENCH_ONLY=query_warm LITCHI_ODS_BENCH_SAMPLES=2 run valgrind --tool=callgrind \
    --toggle-collect='*unified_query*' --callgrind-out-file="$OUT/cg-ods-$tag.out" "$dir/ods_unified_bench" > /dev/null
  run valgrind --tool=callgrind --toggle-collect='*Matches*Iterator*next*' \
    --callgrind-out-file="$OUT/cg-odt-$tag.out" "$dir/odt_paired_bench" > /dev/null
  run valgrind --tool=callgrind --callgrind-out-file="$OUT/cg-rtf-$tag.out" "$dir/rtf_destination_bench" > /dev/null
done
