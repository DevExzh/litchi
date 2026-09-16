#!/usr/bin/env bash
# capture_perf.sh <probe-binary> <leg-name> <out-dir>
#
# Native hardware counters, isolated the same way: difference a large-sample
# and a small-sample child and divide by the extra operations. Change 0579
# showed instruction share mis-ranks pointer-chase work, and callgrind's
# software SHA-256 inflates every fingerprint share, so cycles are taken here
# natively beside the callgrind instruction counts.
set -euo pipefail
BIN=$(readlink -f "$1")
LEG="$2"
mkdir -p "$3"; OUT=$(readlink -f "$3")
REPO=${REPO:-/home/zhuhe/code/litchi}
CPU=${CPU:-9}
SCRATCH=${SCRATCH:-/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0633}
export TMPDIR="$SCRATCH/tmp" RAYON_NUM_THREADS=1 OMP_NUM_THREADS=1
mkdir -p "$TMPDIR"

# stem:path:warmups:small:large
fixtures=(
  "54016:$REPO/test-data/poi/test-data/spreadsheet/54016.xls:3:10:40"
  "cv:$REPO/test-data/ole/xls/WithCustomViews.xls:5:20:120"
  "formula:$REPO/test-data/ole/xls/FormulaEvalTestData.xls:5:20:120"
)

operations_for() {
  case "$1" in
    formula) echo "open number-generic string-generic noop-generic" ;;
    *) echo "open number-plan number-source-backed number-generic string-generic noop-generic" ;;
  esac
}

count() {
  local stem=$1 path=$2 warmups=$3 op=$4 samples=$5
  setarch x86_64 -R taskset -c "$CPU" \
    perf stat -x, -e cycles,instructions,branches,branch-misses,task-clock \
    -o "$OUT/perf-$LEG-$stem-$op-s$samples.csv" \
    "$BIN" --input "$path" --operation "$op" \
    --warmups "$warmups" --samples "$samples" \
    > /dev/null 2>/dev/null
  echo "counted $LEG/$stem/$op/$samples"
}

: > "$OUT/perf-pairs-$LEG.txt"
for fixture in "${fixtures[@]}"; do
  IFS=: read -r stem path warmups small large <<<"$fixture"
  for op in $(operations_for "$stem"); do
    count "$stem" "$path" "$warmups" "$op" "$small"
    count "$stem" "$path" "$warmups" "$op" "$large"
    echo "$stem $op $small $large" >> "$OUT/perf-pairs-$LEG.txt"
  done
done
