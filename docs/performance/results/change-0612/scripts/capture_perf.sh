#!/usr/bin/env bash
# capture_perf.sh <leg-name> <out-dir>
#
# Native hardware counters, isolated the same way: difference a large-sample
# and a small-sample child and divide by the extra operations.  Change 0579
# showed instruction share mis-ranks pointer-chase work and change 0604 showed
# callgrind's per-byte `rep movsb`/`rep stosb` accounting overstates every bulk
# copy, so the byte term this change moves is priced natively as well as under
# callgrind.  Five repetitions per cell so a median can be taken before the
# pair is differenced.
set -euo pipefail
LEG="$1"
mkdir -p "$2"; OUT=$(readlink -f "$2")
REPO=${REPO:-/home/zhuhe/code/litchi}
BINDIR=${BINDIR:?BINDIR must point at the leg binaries}
CPU=${CPU:-8}
SCRATCH=${SCRATCH:?SCRATCH must be set}
export TMPDIR="$SCRATCH/tmp" RAYON_NUM_THREADS=1 OMP_NUM_THREADS=1
mkdir -p "$TMPDIR"
BIN="$BINDIR/leg-$LEG"
OPS=${OPS:-"open one-cell"}
REPS=${REPS:-5}
MODE=${MODE:-owned-readat}

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
          "$BIN" --input "$path" --mode "$MODE" --operation "$op" \
          --worksheet-index "$sheet" --warmups 20 --samples "$samples" \
          > /dev/null 2>/dev/null
      done
      echo "counted $LEG/$stem/$op/$samples x$REPS"
    done
  done
done
