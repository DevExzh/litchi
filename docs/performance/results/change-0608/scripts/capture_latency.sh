#!/usr/bin/env bash
# capture_latency.sh <out-dir> <leg-a-binary> <leg-b-binary> <b-name>
#
# Paired wall clock in A1 B1 B2 A2 order, as changes 0576, 0595 and 0604 took
# it: four rounds per cell, the two inner rounds the B leg.  a2 against a1 and
# b2 against b1 are the same binary against itself in the same window and are
# the measured floor.  Neither B leg here is a candidate: both are change
# 0608's measurement scaffolds, timed to say whether the ceiling they bound is
# separable from this host's noise at all.
set -euo pipefail
mkdir -p "$1"; OUT=$(readlink -f "$1")
A=$(readlink -f "$2"); B=$(readlink -f "$3"); BNAME="$4"
REPO=${REPO:-/home/zhuhe/code/litchi}
CPU=${CPU:-28}
SCRATCH=${SCRATCH:-/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0608}
export TMPDIR="$SCRATCH/tmp" RAYON_NUM_THREADS=1 OMP_NUM_THREADS=1
mkdir -p "$TMPDIR"
SAMPLES=${SAMPLES:-400}
WARMUPS=${WARMUPS:-50}

cells=(
  "flagship:$REPO/test-data/ole/xls/ConditionalFormattingSamples.xls:1"
  "cv:$REPO/test-data/ole/xls/WithCustomViews.xls:1"
  "54016:$REPO/test-data/poi/test-data/spreadsheet/54016.xls:0"
)
uptime > "$OUT/quiescence-$BNAME.log"
for round in "a1:$A" "b1:$B" "b2:$B" "a2:$A"; do
  name=${round%%:*}; bin=${round#*:}
  for cell in "${cells[@]}"; do
    IFS=: read -r stem path sheet <<<"$cell"
    for mode in owned-readat file-source; do
      setarch x86_64 -R taskset -c "$CPU" \
        "$bin" --input "$path" --mode "$mode" --operation open \
        --worksheet-index "$sheet" --warmups "$WARMUPS" --samples "$SAMPLES" \
        > "$OUT/lat-$BNAME-$name-$stem-$mode.json" 2>/dev/null
    done
  done
  echo "round $name done"
done
uptime >> "$OUT/quiescence-$BNAME.log"
