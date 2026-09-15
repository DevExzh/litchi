#!/usr/bin/env bash
# capture_latency.sh <before-binary> <after-binary> <out-dir>
#
# A/B/B/A wall-clock capture. Each leg is a separate child; the two directions
# bracket each other, so monotonic drift in the host shows up as a disagreement
# between them, and a1 against a2 (and b1 against b2) is the same binary against
# itself in the same window -- the measured floor.
#
# `open`, `list` and `one-cell` exist on both legs: they are change 0605's
# no-regression evidence, and the a1/a2 pair states the floor they are judged
# against. The whole-sheet scenarios exist only on the after leg, so their b1/b2
# rounds state their own floor and their paired comparison is scan against
# per-cell inside one binary.
#
# `54016.xls` has one worksheet, so its per-sheet operations use
# `--worksheet-index 0`; the flagship's cell-bearing sheets are 1 (a query) and
# 11 (the widest walk); `WithCustomViews.xls` stores its cells on sheet 0.
set -euo pipefail
BEFORE=$(readlink -f "$1")
AFTER=$(readlink -f "$2")
mkdir -p "$3"; OUT=$(readlink -f "$3")
REPO=${REPO:-/home/zhuhe/code/litchi}
CPU=${CPU:-25}
SCRATCH=${SCRATCH:-/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0605}
export TMPDIR="$SCRATCH/tmp" RAYON_NUM_THREADS=1 OMP_NUM_THREADS=1
mkdir -p "$TMPDIR"

# stem:path:query-sheet:walk-sheet
cells=(
  "flagship:$REPO/test-data/ole/xls/ConditionalFormattingSamples.xls:1:11"
  "cv:$REPO/test-data/ole/xls/WithCustomViews.xls:1:0"
  "54016:$REPO/test-data/poi/test-data/spreadsheet/54016.xls:0:0"
)

for round in a1 b1 b2 a2; do
  case "$round" in
    a1|a2) BIN="$BEFORE"; NEW=0 ;;
    b1|b2) BIN="$AFTER";  NEW=1 ;;
  esac
  for cell in "${cells[@]}"; do
    IFS=: read -r stem path sheet walk <<<"$cell"
    mkdir -p "$OUT/$round/$stem"
    for mode in owned-readat file-source; do
      for op in open list one-cell; do
        setarch x86_64 -R taskset -c "$CPU" "$BIN" \
          --input "$path" --mode "$mode" --operation "$op" \
          --worksheet-index "$sheet" --warmups "${W:-20}" --samples "${S:-60}" \
          > "$OUT/$round/$stem/$mode-$op.json" 2>/dev/null
      done
      if [ "$NEW" = 1 ]; then
        setarch x86_64 -R taskset -c "$CPU" "$BIN" \
          --input "$path" --mode "$mode" --operation second-cell \
          --worksheet-index "$sheet" --warmups "${W:-20}" --samples "${S:-60}" \
          > "$OUT/$round/$stem/$mode-second-cell.json" 2>/dev/null
        setarch x86_64 -R taskset -c "$CPU" "$BIN" \
          --input "$path" --mode "$mode" --operation full-text \
          --warmups "${WT:-5}" --samples "${ST:-30}" \
          > "$OUT/$round/$stem/$mode-full-text.json" 2>/dev/null
        setarch x86_64 -R taskset -c "$CPU" "$BIN" \
          --input "$path" --mode "$mode" --operation all-cells \
          --all-cells-strategy scan --worksheet-index "$walk" \
          --warmups "${WW:-5}" --samples "${SW:-30}" \
          > "$OUT/$round/$stem/$mode-all-cells-scan.json" 2>/dev/null
        setarch x86_64 -R taskset -c "$CPU" "$BIN" \
          --input "$path" --mode "$mode" --operation all-cells \
          --all-cells-strategy per-cell --per-cell-limit 256 \
          --worksheet-index "$walk" \
          --warmups "${WW:-5}" --samples "${SW:-30}" \
          > "$OUT/$round/$stem/$mode-all-cells-per-cell.json" 2>/dev/null
      fi
    done
    echo "captured $round/$stem"
  done
done
