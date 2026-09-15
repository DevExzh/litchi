#!/usr/bin/env bash
# Callgrind isolation pairs and Deflate-construction counts for change 0618.
# Callgrind counts `rep stos` once per byte, so the compressor's 300 KiB state
# zeroing is over-priced here; `run-perf.sh` carries the native price.
set -u
S="$1"; CPU=12
CFS=/home/zhuhe/code/litchi-worktrees/before-8fe9efa55/test-data/ooxml/xlsx/ConditionalFormattingSamples.xlsx
run() { local tag="$1"; shift
  taskset -c $CPU valgrind --tool=callgrind --callgrind-out-file="$S/cg/$tag.out" --cache-sim=no --branch-sim=no "$@" >/dev/null 2>"$S/cg/$tag.err"
  echo -e "$tag\t$(grep -oP 'Collected\s*:\s*\K[0-9]+' "$S/cg/$tag.err" | tail -1)"
}
for leg in before after; do
  OPC=/home/zhuhe/code/litchi-worktrees/targets/0618-opc-$leg/release/probe0618-opc
  PPTX=/home/zhuhe/code/litchi-worktrees/targets/0618-p0607-$leg/release/probe0618-pptx
  for sc in addrel addrelN:8 addrelall; do
    tag=$(echo $sc | tr ':' '-')
    run "$leg-opc-$tag-4"  $OPC savebench $CFS $sc 4
    run "$leg-opc-$tag-24" $OPC savebench $CFS $sc 24
  done
  run "$leg-pptx-create50-1"  $PPTX create 50 1
  run "$leg-pptx-create50-6"  $PPTX create 50 6
  run "$leg-pptx-create1-1"   $PPTX create 1 1
  run "$leg-pptx-create1-21"  $PPTX create 1 21
done
