#!/bin/sh
# Record 0317 wall-time driver. usage: measure.sh BIN_DIR OUT_DIR
# BIN_DIR holds stripped `profiling` builds {docx,pptx}_large_bench-{base,after}
# (SHA-256 in binaries.sha256) of
#   cargo build --profile profiling -p litchi-ops \
#     --features docx,doc,odp,xls,rtf,odt,pptx,ppt --example <name>
# base: this branch with crates/{xml-minifier,litchi-opc,litchi-docx}/src at
# trunk 30eed270b5 (so it has the new pptx_large_bench); after: this branch.
# Processes run B A A B on one pinned CPU.
set -eu
D=$1 OUT=$2 CPU=${CPU:-30}
mkdir -p "$OUT"
for b in docx pptx; do
  for leg in base-1 after-1 after-2 base-2; do
    bin=$D/${b}_large_bench-${leg%-*}
    if [ $b = pptx ]; then PPTX_LARGE_LOOPS=5 taskset -c "$CPU" "$bin" > "$OUT/$b-large-$leg.txt"
    else taskset -c "$CPU" "$bin" > "$OUT/$b-large-$leg.txt"; fi
  done
done
