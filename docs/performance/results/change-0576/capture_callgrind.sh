#!/usr/bin/env bash
# capture_callgrind.sh <binary> <leg-name> <out-dir>
#
# Instruction attribution by the isolation method change 0574 used: run the same
# child at a small and a large sample count, difference the two profiles, and
# divide by the number of extra operations. Warmups are identical in both legs so
# that everything outside the measured operation cancels.
set -euo pipefail
BIN=$(readlink -f "$1")
LEG="$2"
mkdir -p "$3"; OUT=$(readlink -f "$3")
REPO=${REPO:-/home/zhuhe/code/litchi}
CPU=${CPU:-17}
SCRATCH=${SCRATCH:-/tmp/claude-1001/-home-zhuhe-code-litchi/02240c68-6724-42ba-929e-c4157529483d/scratchpad/change-0576}
export TMPDIR="$SCRATCH/tmp" RAYON_NUM_THREADS=1 OMP_NUM_THREADS=1
mkdir -p "$TMPDIR"

# stem:fixture:small-samples:large-samples
cells=(
  "flagship:$REPO/test-data/ole/xls/ConditionalFormattingSamples.xls:20:220"
  "cv:$REPO/test-data/ole/xls/WithCustomViews.xls:20:120"
  "54016:$REPO/test-data/poi/test-data/spreadsheet/54016.xls:10:60"
)

for cell in "${cells[@]}"; do
  IFS=: read -r stem path small large <<<"$cell"
  for pair in "small:$small" "large:$large"; do
    label=${pair%%:*}; samples=${pair#*:}
    raw="$OUT/cg-$LEG-$stem-s$label.out"
    setarch x86_64 -R taskset -c "$CPU" \
      valgrind --tool=callgrind --callgrind-out-file="$raw" --quiet \
      "$BIN" --input "$path" --mode owned-readat --operation open \
      --warmups 2 --samples "$samples" > /dev/null 2> "$OUT/cg-$LEG-$stem-s$label.stderr"
    callgrind_annotate --threshold=99.9 "$raw" > "$OUT/ann-$LEG-$stem-s$label.txt"
    callgrind_annotate --threshold=99.9 --inclusive=yes "$raw" > "$OUT/incl-$LEG-$stem-s$label.txt"
    rm -f "$raw"
    echo "annotated $LEG/$stem/$label"
  done
done
