#!/usr/bin/env bash
# capture_perf.sh <binary> <leg-name> <out-dir>
#
# Hardware counters, isolated the way change 0574 isolated them: difference an
# 1100-sample and a 100-sample child and divide by 1000.
set -euo pipefail
BIN=$(readlink -f "$1")
LEG="$2"
mkdir -p "$3"; OUT=$(readlink -f "$3")
REPO=${REPO:-/home/zhuhe/code/litchi}
CPU=${CPU:-17}
SCRATCH=${SCRATCH:-/tmp/claude-1001/-home-zhuhe-code-litchi/02240c68-6724-42ba-929e-c4157529483d/scratchpad/change-0576}
export TMPDIR="$SCRATCH/tmp" RAYON_NUM_THREADS=1 OMP_NUM_THREADS=1
mkdir -p "$TMPDIR"

fixtures=(
  "flagship:$REPO/test-data/ole/xls/ConditionalFormattingSamples.xls"
  "cv:$REPO/test-data/ole/xls/WithCustomViews.xls"
  "54016:$REPO/test-data/poi/test-data/spreadsheet/54016.xls"
)

for entry in "${fixtures[@]}"; do
  stem=${entry%%:*}
  path=${entry#*:}
  for samples in 100 1100; do
    setarch x86_64 -R taskset -c "$CPU" \
      perf stat -x, -e instructions,branches,branch-misses,task-clock \
      -o "$OUT/perf-$LEG-$stem-s$samples.csv" \
      "$BIN" --input "$path" --mode owned-readat --operation open \
      --warmups 20 --samples "$samples" > /dev/null 2>/dev/null
    echo "counted $LEG/$stem/$samples"
  done
done
