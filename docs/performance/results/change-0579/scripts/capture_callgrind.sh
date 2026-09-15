#!/usr/bin/env bash
# capture_callgrind.sh <attribution-binary> <out-dir> <tag>
# Callgrind isolation pairs, change 0574's method: a small-sample and a
# large-sample child, differenced and divided by the sample delta.
set -euo pipefail
ATTR=$(readlink -f "$1"); mkdir -p "$2"; OUT=$(readlink -f "$2"); TAG=$3
CPU=${CPU:-17}
SCRATCH=/tmp/claude-1001/-home-zhuhe-code-litchi/02240c68-6724-42ba-929e-c4157529483d/scratchpad/change-0579
export TMPDIR="$SCRATCH/tmp" RAYON_NUM_THREADS=1 OMP_NUM_THREADS=1
mkdir -p "$TMPDIR" "$OUT/cg"
run () { # stem input samples
  local stem=$1 input=$2 samples=$3
  local out="$OUT/cg/${TAG}-${stem}-s${samples}.out"
  setarch x86_64 -R taskset -c "$CPU" \
    valgrind --tool=callgrind --callgrind-out-file="$out" --cache-sim=no --branch-sim=no \
    "$ATTR" --input "$input" --mode owned-readat --operation open \
    --warmups 1 --samples "$samples" > /dev/null 2>"$OUT/cg/${TAG}-${stem}-s${samples}.stderr"
  callgrind_annotate --threshold=99 "$out"            > "$OUT/ann-${TAG}-${stem}-s${samples}.txt"
  callgrind_annotate --threshold=99 --inclusive=yes "$out" > "$OUT/incl-${TAG}-${stem}-s${samples}.txt"
  callgrind_annotate --threshold=99 --tree=caller "$out"   > "$OUT/tree-${TAG}-${stem}-s${samples}.txt"
  echo "wrote ${TAG}-${stem}-s${samples}"
}
FLAG=/home/zhuhe/code/litchi/test-data/ole/xls/ConditionalFormattingSamples.xls
CV=/home/zhuhe/code/litchi/test-data/ole/xls/WithCustomViews.xls
K54016=/home/zhuhe/code/litchi/test-data/poi/test-data/spreadsheet/54016.xls
run flagship "$FLAG" 20
run flagship "$FLAG" 220
run cv "$CV" 20
run cv "$CV" 120
run 54016 "$K54016" 10
run 54016 "$K54016" 60
