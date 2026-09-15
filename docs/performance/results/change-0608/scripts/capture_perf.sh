#!/usr/bin/env bash
# capture_perf.sh <binary> <leg-name> <out-dir>
#
# Native hardware counters, isolated the way change 0574 isolated them and
# changes 0584, 0595 and 0604 reused: difference a large-sample and a
# small-sample child and divide by the extra operations.  Change 0579 showed
# that instruction share mis-ranks pointer-chase work, and change 0595's own
# record shows callgrind and perf disagree in level on this very path, so the
# SST walk is priced natively as well as under callgrind.
set -euo pipefail
BIN=$(readlink -f "$1")
LEG="$2"
mkdir -p "$3"; OUT=$(readlink -f "$3")
REPO=${REPO:-/home/zhuhe/code/litchi}
CPU=${CPU:-28}
SCRATCH=${SCRATCH:-/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0608}
export TMPDIR="$SCRATCH/tmp" RAYON_NUM_THREADS=1 OMP_NUM_THREADS=1
mkdir -p "$TMPDIR"
OPS=${OPS:-"open one-cell"}
REPS=${REPS:-5}

# stem:path:worksheet-index:small:large
cells=(
  "flagship:$REPO/test-data/ole/xls/ConditionalFormattingSamples.xls:1:100:1100"
  "cv:$REPO/test-data/ole/xls/WithCustomViews.xls:1:100:1100"
  "54016:$REPO/test-data/poi/test-data/spreadsheet/54016.xls:0:100:1100"
)

for cell in "${cells[@]}"; do
  IFS=: read -r stem path sheet small large <<<"$cell"
  for op in $OPS; do
    for samples in "$small" "$large"; do
      for r in $(seq 1 "$REPS"); do
        setarch x86_64 -R taskset -c "$CPU" \
          perf stat -x, -e cycles,instructions,branches,branch-misses,task-clock \
          -o "$OUT/perf-$LEG-$stem-$op-s$samples-r$r.csv" \
          "$BIN" --input "$path" --mode owned-readat --operation "$op" \
          --worksheet-index "$sheet" --warmups 20 --samples "$samples" \
          > /dev/null 2>/dev/null
      done
      echo "counted $LEG/$stem/$op/$samples x$REPS"
    done
  done
done
