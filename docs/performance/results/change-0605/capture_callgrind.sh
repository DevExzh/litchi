#!/usr/bin/env bash
# capture_callgrind.sh <binary> <leg-name> <out-dir>
#
# Instruction attribution by change 0574's isolation method, as changes 0576,
# 0584 and 0595 used it: run the same child at a small and a large sample count,
# difference the two profiles and divide by the extra operations. Warmups are
# identical in both legs so everything outside the measured operation cancels.
#
# `open` and `one-cell` run on both legs and are change 0605's no-regression
# evidence. `all-cells` runs on the after leg only, in both strategies, because
# the whole-sheet walk does not exist on the before leg; the two strategies are
# the paired legs of that scenario.
set -euo pipefail
BIN=$(readlink -f "$1")
LEG="$2"
mkdir -p "$3"; OUT=$(readlink -f "$3")
REPO=${REPO:-/home/zhuhe/code/litchi}
CPU=${CPU:-25}
SCRATCH=${SCRATCH:-/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0605}
export TMPDIR="$SCRATCH/tmp" RAYON_NUM_THREADS=1 OMP_NUM_THREADS=1
mkdir -p "$TMPDIR"

# stem:path:worksheet-index:open-small:open-large:cell-small:cell-large
cells=(
  "flagship:$REPO/test-data/ole/xls/ConditionalFormattingSamples.xls:1:20:220:20:220"
  "cv:$REPO/test-data/ole/xls/WithCustomViews.xls:1:20:120:20:120"
  "54016:$REPO/test-data/poi/test-data/spreadsheet/54016.xls:0:10:60:5:25"
)

run_pair() {
  local stem=$1 path=$2 sheet=$3 op=$4 label=$5 small=$6 large=$7
  shift 7
  for pair in "small:$small" "large:$large"; do
    local size=${pair%%:*} samples=${pair#*:}
    local raw="$OUT/cg-$LEG-$stem-$label-s$size.out"
    setarch x86_64 -R taskset -c "$CPU" \
      valgrind --tool=callgrind --callgrind-out-file="$raw" --quiet \
      --cache-sim=no --branch-sim=no \
      "$BIN" --input "$path" --mode owned-readat --operation "$op" "$@" \
      --worksheet-index "$sheet" --warmups 2 --samples "$samples" \
      > /dev/null 2> "$OUT/cg-$LEG-$stem-$label-s$size.stderr"
    callgrind_annotate --threshold=99.9 "$raw" > "$OUT/ann-$LEG-$stem-$label-s$size.txt"
    rm -f "$raw"
    echo "annotated $LEG/$stem/$label/$size"
  done
}

for cell in "${cells[@]}"; do
  IFS=: read -r stem path sheet osmall olarge csmall clarge <<<"$cell"
  run_pair "$stem" "$path" "$sheet" open open "$osmall" "$olarge"
  run_pair "$stem" "$path" "$sheet" one-cell one-cell "$csmall" "$clarge"
done

# The whole-sheet walk exists only on the after leg.
if [ "$LEG" = after ]; then
  # stem:path:sheet:small:large  (a walk is far heavier than one query)
  walks=(
    "flagship:$REPO/test-data/ole/xls/ConditionalFormattingSamples.xls:1:5:25"
    "cv:$REPO/test-data/ole/xls/WithCustomViews.xls:0:2:6"
    "54016:$REPO/test-data/poi/test-data/spreadsheet/54016.xls:0:2:6"
  )
  for cell in "${walks[@]}"; do
    IFS=: read -r stem path sheet small large <<<"$cell"
    run_pair "$stem" "$path" "$sheet" all-cells all-cells-scan "$small" "$large" \
      --all-cells-strategy scan
    run_pair "$stem" "$path" "$sheet" all-cells all-cells-per-cell "$small" "$large" \
      --all-cells-strategy per-cell --per-cell-limit 64
  done
fi
