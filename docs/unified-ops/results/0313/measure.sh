#!/bin/sh
# Record 0313 wall-time driver. BASE and AFTER are stripped `profiling` builds
# (SHA-256 in binaries.sha256) of
#   cargo build --profile profiling -p litchi-ops \
#     --features docx,doc,odp,xls,rtf,odt,pptx,ppt --example docx_large_bench
# BASE: this branch with crates/{xml-minifier,litchi-opc}/src at trunk 12724da2d5
# (so it has the new fresh_* legs); AFTER: this branch.
# Processes run B A A B on one pinned CPU; the harness times its legs one
# after another (two warm-ups, nine samples of 20 calls, the median).
set -eu
BASE=$1 AFTER=$2 OUT=$3 CPU=${CPU:-30}
mkdir -p "$OUT"
for leg in base-1 after-1 after-2 base-2; do
  bin=$BASE; case $leg in after*) bin=$AFTER;; esac
  taskset -c "$CPU" "$bin" > "$OUT/docx-large-$leg.txt"
done
