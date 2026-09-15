#!/usr/bin/env bash
# Native isolation pairs for change 0618: `perf stat` prices what callgrind
# over-counts (the compressor's 300 KiB state zeroing is one `rep stos` here,
# one instruction per byte there).
set -u
S="$1"; CPU=12
CFS=/home/zhuhe/code/litchi-worktrees/before-8fe9efa55/test-data/ooxml/xlsx/ConditionalFormattingSamples.xlsx
row() { # row <leg> <case> <binary> <argv...>
  local leg="$1" case="$2"; shift 2
  echo -n "$leg	$case	"
  taskset -c $CPU perf stat -x, -r 5 -e cycles,instructions,page-faults "$@" 2>&1 >/dev/null \
    | awk -F, '{printf "%s=%s\t", $3, $1}'
  echo
}
for leg in before after; do
  case $leg in before) d=before;; after) d=after;; esac
  OPC=/home/zhuhe/code/litchi-worktrees/targets/0618-opc-$d/release/probe0618-opc
  PPTX=/home/zhuhe/code/litchi-worktrees/targets/0618-p0607-$d/release/probe0618-pptx
  for sc in addrelall addrelN:8 addrel; do
    row "$leg" "opc-$sc-200"  $OPC savebench $CFS $sc 200
    row "$leg" "opc-$sc-1200" $OPC savebench $CFS $sc 1200
  done
  row "$leg" "pptx-create50-5"   $PPTX create 50 5
  row "$leg" "pptx-create50-55"  $PPTX create 50 55
  row "$leg" "pptx-create1-20"   $PPTX create 1 20
  row "$leg" "pptx-create1-220"  $PPTX create 1 220
done
