#!/usr/bin/env bash
# capture.sh <binary> <outdir> <tag> <stem> <input> <operation> <small> <large> <cpu> [extra args...]
set -euo pipefail
BIN=$1; OUT=$2; TAG=$3; stem=$4; input=$5; op=$6; small=$7; large=$8; cpu=$9; shift 9
SP=/tmp/claude-1001/-home-zhuhe-code-litchi/14c44904-927d-4351-97ac-5611bafb5316/scratchpad
export TMPDIR="$SP/tmp" RAYON_NUM_THREADS=1 OMP_NUM_THREADS=1
mkdir -p "$TMPDIR" "$OUT/cg" "$OUT/ann"
for samples in "$small" "$large"; do
  o="$OUT/cg/${TAG}-${stem}-${op}-s${samples}.out"
  setarch x86_64 -R taskset -c "$cpu" \
    valgrind --tool=callgrind --callgrind-out-file="$o" --cache-sim=no --branch-sim=no \
      "$BIN" --input "$input" --mode owned-readat --operation "$op" \
      --warmups 1 --samples "$samples" "$@" \
      > /dev/null 2>"$OUT/cg/${TAG}-${stem}-${op}-s${samples}.stderr"
  callgrind_annotate --threshold=99.9 "$o"               > "$OUT/ann/self-${TAG}-${stem}-${op}-s${samples}.txt"
  callgrind_annotate --threshold=99.9 --inclusive=yes "$o" > "$OUT/ann/incl-${TAG}-${stem}-${op}-s${samples}.txt"
  callgrind_annotate --threshold=99.5 --tree=caller "$o"  > "$OUT/ann/tree-${TAG}-${stem}-${op}-s${samples}.txt"
done
echo "done ${stem}-${op}"
