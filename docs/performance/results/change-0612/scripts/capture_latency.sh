#!/usr/bin/env bash
# capture_latency.sh <out-dir> <leg-a-binary> <leg-b-binary> <b-name> [ops]
#
# Paired wall clock in A1 B1 B2 A2 order, as changes 0576, 0595, 0604 and 0608
# took it: four rounds per cell, the two inner rounds the B leg.  a2 against a1
# and b2 against b1 are the same binary against itself in the same window and
# are the measured floor.
#
# Both source modes are timed for every cell, because this change trades read
# calls for read bytes and the two modes price a read call differently:
# `owned-readat` is a bounds check and a memcpy, `file-source` is a `pread64`
# plus a `statx` freshness observation.
set -euo pipefail
mkdir -p "$1"; OUT=$(readlink -f "$1")
A=$(readlink -f "$2"); B=$(readlink -f "$3"); BNAME="$4"
OPS=${5:-open}
REPO=${REPO:-/home/zhuhe/code/litchi}
CPU=${CPU:-8}
SCRATCH=${SCRATCH:?SCRATCH must be set}
export TMPDIR="$SCRATCH/tmp" RAYON_NUM_THREADS=1 OMP_NUM_THREADS=1
mkdir -p "$TMPDIR"
SAMPLES=${SAMPLES:-400}
WARMUPS=${WARMUPS:-50}

cells=(
  "flagship:$REPO/test-data/ole/xls/ConditionalFormattingSamples.xls:1"
  "cv:$REPO/test-data/ole/xls/WithCustomViews.xls:1"
  "54016:$REPO/test-data/poi/test-data/spreadsheet/54016.xls:0"
)
uptime > "$OUT/quiescence-$BNAME.log"
for round in "a1:$A" "b1:$B" "b2:$B" "a2:$A"; do
  name=${round%%:*}; bin=${round#*:}
  for cell in "${cells[@]}"; do
    IFS=: read -r stem path sheet <<<"$cell"
    for mode in owned-readat file-source; do
      for op in $OPS; do
        setarch x86_64 -R taskset -c "$CPU" \
          "$bin" --input "$path" --mode "$mode" --operation "$op" \
          --worksheet-index "$sheet" --warmups "$WARMUPS" --samples "$SAMPLES" \
          > "$OUT/lat-$BNAME-$name-$stem-$mode-$op.json" 2>/dev/null
      done
    done
  done
  echo "round $name done"
done
uptime >> "$OUT/quiescence-$BNAME.log"
