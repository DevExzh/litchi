#!/usr/bin/env bash
# capture_counters.sh <attribution-binary> <leg-name> <out-dir>
#
# Deterministic logical counters (reads, read bytes, version observations, len
# observations, seeks) for the three fixtures this change measures, over the two
# in-memory source modes and all three operations. These are the oracle for
# "not one byte of I/O moved": change 0595 removes CPU work only.
#
# `54016.xls` has one worksheet, so its `one-cell` operation is driven with
# `--worksheet-index 0`; the other two fixtures keep the harness default of 1.
set -euo pipefail
ATTR=$(readlink -f "$1")
LEG="$2"
mkdir -p "$3"; OUT=$(readlink -f "$3")
REPO=${REPO:-/home/zhuhe/code/litchi}
CPU=${CPU:-15}
SCRATCH=${SCRATCH:-/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0595}
export TMPDIR="$SCRATCH/tmp" RAYON_NUM_THREADS=1 OMP_NUM_THREADS=1
mkdir -p "$TMPDIR"

# stem:path:worksheet-index
cells=(
  "flagship:$REPO/test-data/ole/xls/ConditionalFormattingSamples.xls:1"
  "cv:$REPO/test-data/ole/xls/WithCustomViews.xls:1"
  "54016:$REPO/test-data/poi/test-data/spreadsheet/54016.xls:0"
)

for cell in "${cells[@]}"; do
  IFS=: read -r stem path sheet <<<"$cell"
  for mode in owned-readat file-source; do
    for op in open list one-cell; do
      setarch x86_64 -R taskset -c "$CPU" "$ATTR" \
        --input "$path" --mode "$mode" --operation "$op" \
        --worksheet-index "$sheet" --warmups "${W:-20}" --samples "${S:-100}" \
        > "$OUT/$LEG-$stem-$mode-$op.json" 2> "$OUT/$LEG-$stem-$mode-$op.stderr"
      echo "wrote $OUT/$LEG-$stem-$mode-$op.json"
    done
  done
done
