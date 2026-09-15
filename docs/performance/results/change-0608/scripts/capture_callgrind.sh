#!/usr/bin/env bash
# capture_callgrind.sh <binary> <leg-name> <out-dir>
#
# Change 0574's isolation method, as changes 0576, 0584 and 0595 used it: run
# the same child at a small and a large sample count, difference the two
# profiles and divide by the extra operations, so everything that runs once per
# child cancels.  Both a self-cost and an inclusive annotation are taken from
# the same raw profile, because change 0608 needs the *inclusive* cost of
# `scan_shared_string_records` (its share of the open) as well as the self costs
# of the symbols inside it.
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

# stem:path:worksheet-index:open-small:open-large:cell-small:cell-large
cells=(
  "flagship:$REPO/test-data/ole/xls/ConditionalFormattingSamples.xls:1:20:220:20:220"
  "cv:$REPO/test-data/ole/xls/WithCustomViews.xls:1:20:120:20:120"
  "54016:$REPO/test-data/poi/test-data/spreadsheet/54016.xls:0:10:60:5:25"
)

for cell in "${cells[@]}"; do
  IFS=: read -r stem path sheet osmall olarge csmall clarge <<<"$cell"
  for op in $OPS; do
    if [ "$op" = open ]; then small=$osmall; large=$olarge; else small=$csmall; large=$clarge; fi
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
