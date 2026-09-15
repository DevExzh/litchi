#!/usr/bin/env bash
# capture_latency.sh <before-binary> <after-binary> <out-dir>
#
# A/B/B/A wall-clock capture. Each leg is a separate child; the two directions
# bracket each other, so monotonic drift in the host shows up as a disagreement
# between them, and a1 against a2 (and b1 against b2) is the same binary against
# itself in the same window -- the measured floor.
#
# `54016.xls` has one worksheet, so its one-cell operation is driven with
# `--worksheet-index 0`.
set -euo pipefail
BEFORE=$(readlink -f "$1")
AFTER=$(readlink -f "$2")
mkdir -p "$3"; OUT=$(readlink -f "$3")
REPO=${REPO:-/home/zhuhe/code/litchi}
CPU=${CPU:-15}
SCRATCH=${SCRATCH:-/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0595}
export TMPDIR="$SCRATCH/tmp" RAYON_NUM_THREADS=1 OMP_NUM_THREADS=1
mkdir -p "$TMPDIR"

cells=(
  "flagship:$REPO/test-data/ole/xls/ConditionalFormattingSamples.xls:1"
  "cv:$REPO/test-data/ole/xls/WithCustomViews.xls:1"
  "54016:$REPO/test-data/poi/test-data/spreadsheet/54016.xls:0"
)

for round in a1 b1 b2 a2; do
  case "$round" in
    a1|a2) BIN="$BEFORE" ;;
    b1|b2) BIN="$AFTER" ;;
  esac
  for cell in "${cells[@]}"; do
    IFS=: read -r stem path sheet <<<"$cell"
    mkdir -p "$OUT/$round/$stem"
    for mode in owned-readat file-source; do
      for op in open list one-cell; do
        setarch x86_64 -R taskset -c "$CPU" "$BIN" \
          --input "$path" --mode "$mode" --operation "$op" \
          --worksheet-index "$sheet" --warmups "${W:-20}" --samples "${S:-60}" \
          > "$OUT/$round/$stem/$mode-$op.json" 2>/dev/null
      done
    done
    echo "captured $round/$stem"
  done
done
