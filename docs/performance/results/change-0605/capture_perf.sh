#!/usr/bin/env bash
# capture_perf.sh <binary> <leg-name> <out-dir>
#
# Hardware counters, isolated the way change 0574 isolated them: difference a
# large-sample and a small-sample child and divide by the extra operations.
# Change 0579 showed instruction share mis-ranks pointer-chase work, so cycles
# and IPC are taken natively beside the callgrind instruction counts.
set -euo pipefail
BIN=$(readlink -f "$1")
LEG="$2"
mkdir -p "$3"; OUT=$(readlink -f "$3")
REPO=${REPO:-/home/zhuhe/code/litchi}
CPU=${CPU:-25}
SCRATCH=${SCRATCH:-/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0605}
export TMPDIR="$SCRATCH/tmp" RAYON_NUM_THREADS=1 OMP_NUM_THREADS=1
mkdir -p "$TMPDIR"

# stem:path:worksheet-index:small:large
cells=(
  "flagship:$REPO/test-data/ole/xls/ConditionalFormattingSamples.xls:1:100:1100"
  "cv:$REPO/test-data/ole/xls/WithCustomViews.xls:1:100:1100"
  "54016:$REPO/test-data/poi/test-data/spreadsheet/54016.xls:0:100:1100"
)

count() {
  local stem=$1 path=$2 sheet=$3 op=$4 label=$5 samples=$6
  shift 6
  setarch x86_64 -R taskset -c "$CPU" \
    perf stat -x, -e cycles,instructions,branches,branch-misses,task-clock \
    -o "$OUT/perf-$LEG-$stem-$label-s$samples.csv" \
    "$BIN" --input "$path" --mode owned-readat --operation "$op" "$@" \
    --worksheet-index "$sheet" --warmups 20 --samples "$samples" \
    > /dev/null 2>/dev/null
  echo "counted $LEG/$stem/$label/$samples"
}

for cell in "${cells[@]}"; do
  IFS=: read -r stem path sheet small large <<<"$cell"
  for op in open one-cell; do
    count "$stem" "$path" "$sheet" "$op" "$op" "$small"
    count "$stem" "$path" "$sheet" "$op" "$op" "$large"
  done
done

if [ "$LEG" = after ]; then
  walks=(
    "flagship:$REPO/test-data/ole/xls/ConditionalFormattingSamples.xls:1:20:220"
    "cv:$REPO/test-data/ole/xls/WithCustomViews.xls:0:5:55"
    "54016:$REPO/test-data/poi/test-data/spreadsheet/54016.xls:0:5:55"
  )
  for cell in "${walks[@]}"; do
    IFS=: read -r stem path sheet small large <<<"$cell"
    for samples in "$small" "$large"; do
      count "$stem" "$path" "$sheet" all-cells all-cells-scan "$samples" \
        --all-cells-strategy scan
      count "$stem" "$path" "$sheet" all-cells all-cells-per-cell "$samples" \
        --all-cells-strategy per-cell --per-cell-limit 64
    done
  done
fi
