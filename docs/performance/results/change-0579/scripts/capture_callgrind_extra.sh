#!/usr/bin/env bash
# capture_callgrind_extra.sh <binary> <out-dir> <tag> <stem> <input> <small> <large>
set -euo pipefail
ATTR=$(readlink -f "$1"); mkdir -p "$2"; OUT=$(readlink -f "$2"); TAG=$3; stem=$4; input=$5; small=$6; large=$7
CPU=${CPU:-17}
SCRATCH=/tmp/claude-1001/-home-zhuhe-code-litchi/02240c68-6724-42ba-929e-c4157529483d/scratchpad/change-0579
export TMPDIR="$SCRATCH/tmp" RAYON_NUM_THREADS=1 OMP_NUM_THREADS=1
mkdir -p "$TMPDIR" "$OUT/cg"
for samples in "$small" "$large"; do
  out="$OUT/cg/${TAG}-${stem}-s${samples}.out"
  setarch x86_64 -R taskset -c "$CPU" \
    valgrind --tool=callgrind --callgrind-out-file="$out" --cache-sim=no --branch-sim=no \
    "$ATTR" --input "$input" --mode owned-readat --operation open \
    --warmups 1 --samples "$samples" > /dev/null 2>"$OUT/cg/${TAG}-${stem}-s${samples}.stderr"
  callgrind_annotate --threshold=99 "$out" > "$OUT/ann-${TAG}-${stem}-s${samples}.txt"
  callgrind_annotate --threshold=99 --inclusive=yes "$out" > "$OUT/incl-${TAG}-${stem}-s${samples}.txt"
  callgrind_annotate --threshold=99 --tree=caller "$out" > "$OUT/tree-${TAG}-${stem}-s${samples}.txt"
done
echo "wrote ${TAG}-${stem}"
