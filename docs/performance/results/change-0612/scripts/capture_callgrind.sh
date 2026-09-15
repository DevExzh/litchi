#!/usr/bin/env bash
# capture_callgrind.sh <leg-name> <out-dir>
#
# Change 0574's isolation method as changes 0576, 0584, 0595 and 0608 used it:
# run the same child at a small and a large sample count, difference the two
# profiles and divide by the extra operations, so everything that runs once per
# child cancels.  Both a self-cost and an inclusive annotation come from the
# same raw profile, because the size of the globals pass is the *inclusive*
# cost of `parse_globals` and the size of the read term is the inclusive cost
# of `read_stream_range_hinted`.
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

# stem:path:worksheet-index:small:large
cells=(
  "flagship:$REPO/test-data/ole/xls/ConditionalFormattingSamples.xls:1:20:220"
  "cv:$REPO/test-data/ole/xls/WithCustomViews.xls:1:20:120"
  "54016:$REPO/test-data/poi/test-data/spreadsheet/54016.xls:0:10:60"
)

for cell in "${cells[@]}"; do
  IFS=: read -r stem path sheet small large <<<"$cell"
  for op in $OPS; do
    for pair in "small:$small" "large:$large"; do
      label=${pair%%:*}; samples=${pair#*:}
      raw="$TMPDIR/cg-$LEG-$stem-$op-s$label.out"
      setarch x86_64 -R taskset -c "$CPU" \
        valgrind --tool=callgrind --callgrind-out-file="$raw" --quiet \
        --cache-sim=no --branch-sim=no \
        "$BIN" --input "$path" --mode owned-readat --operation "$op" \
        --worksheet-index "$sheet" --warmups 2 --samples "$samples" \
        > /dev/null 2> "$OUT/cg-$LEG-$stem-$op-s$label.stderr"
      callgrind_annotate --threshold=99.9 "$raw" > "$OUT/ann-$LEG-$stem-$op-s$label.txt"
      callgrind_annotate --inclusive=yes --threshold=99.9 "$raw" \
        > "$OUT/inc-$LEG-$stem-$op-s$label.txt"
      rm -f "$raw"
      echo "annotated $LEG/$stem/$op/$label"
    done
  done
done
