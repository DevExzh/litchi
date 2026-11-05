#!/bin/sh
# Record 0311 measurement driver. BASE and AFTER hold `profiling` builds
# (debug info stripped; SHA-256 in binaries.sha256) of
#   cargo build --profile profiling -p litchi-ops \
#     --features docx,doc,odp,xls,rtf,odt,pptx,ppt --example <name>
# BASE: this branch with crates/{litchi-opc,litchi-docx,litchi-doc}/src
# reverted to 1178eb235c (so the harness has the new legs); AFTER: the branch.
# Processes run B A A B on one pinned CPU; office_gaps_bench pairs its legs
# ABBA in process, docx_large_bench times its legs one after another.
set -eu
BASE=$1 AFTER=$2 OUT=$3 CPU=${CPU:-30}
mkdir -p "$OUT"
cd "$(dirname "$0")/../../../../crates/litchi-ops"
for leg in base-1 after-1 after-2 base-2; do
  dir=$BASE; case $leg in after*) dir=$AFTER;; esac
  taskset -c "$CPU" "$dir/docx_large_bench.stripped" > "$OUT/docx-large-$leg.txt"
  for s in doc_text_set doc_package_text_set; do
    OFFICE_GAPS_SCENARIO=$s taskset -c "$CPU" "$dir/office_gaps_bench.stripped"
  done > "$OUT/doc-$leg.jsonl"
done
