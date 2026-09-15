#!/usr/bin/env bash
# capture_walk.sh <after-binary> <out-dir>
#
# The all-cells scenario's paired legs, inside one binary and one build: the
# whole-sheet walk against the per-cell reading it replaces, at three sample
# sizes of positions, on the worksheet of each fixture that actually stores
# cells. `--per-cell-limit` bounds the per-cell leg because every position it
# reads costs a complete validated scan of the worksheet substream.
#
# Also captures full text, which the walk does not change: it already scanned
# each sheet once. Its rows are this change's control.
set -euo pipefail
BIN=$(readlink -f "$1")
mkdir -p "$2"; OUT=$(readlink -f "$2")
REPO=${REPO:-/home/zhuhe/code/litchi}
CPU=${CPU:-25}
SCRATCH=${SCRATCH:-/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0605}
export TMPDIR="$SCRATCH/tmp" RAYON_NUM_THREADS=1 OMP_NUM_THREADS=1
mkdir -p "$TMPDIR"

# stem:path:walk-sheet
cells=(
  "flagship:$REPO/test-data/ole/xls/ConditionalFormattingSamples.xls:11"
  "cv:$REPO/test-data/ole/xls/WithCustomViews.xls:0"
  "54016:$REPO/test-data/poi/test-data/spreadsheet/54016.xls:0"
)

for cell in "${cells[@]}"; do
  IFS=: read -r stem path sheet <<<"$cell"
  for mode in owned-readat file-source; do
    setarch x86_64 -R taskset -c "$CPU" "$BIN" \
      --input "$path" --mode "$mode" --operation all-cells --all-cells-strategy scan \
      --worksheet-index "$sheet" --warmups "${W:-5}" --samples "${S:-20}" \
      > "$OUT/walk-$stem-$mode-scan.json" 2> "$OUT/walk-$stem-$mode-scan.stderr"
    for limit in 8 64 256; do
      setarch x86_64 -R taskset -c "$CPU" "$BIN" \
        --input "$path" --mode "$mode" --operation all-cells \
        --all-cells-strategy per-cell --per-cell-limit "$limit" \
        --worksheet-index "$sheet" --warmups "${W:-5}" --samples "${S:-20}" \
        > "$OUT/walk-$stem-$mode-per-cell-$limit.json" \
        2> "$OUT/walk-$stem-$mode-per-cell-$limit.stderr"
    done
    echo "walked $stem/$mode (sheet $sheet)"
  done
done
