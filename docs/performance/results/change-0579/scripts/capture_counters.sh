#!/usr/bin/env bash
# capture_counters.sh <attribution-binary> <out-dir>
# Deterministic logical counters for the three OLE2 read scenarios, over two
# source implementations, following change 0564/0565/0568's attribution-child
# setting (--warmups 20 --samples 100).
set -euo pipefail
ATTR=$(readlink -f "$1"); mkdir -p "$2"; OUT=$(readlink -f "$2")
CPU=${CPU:-17}
INPUT=${INPUT:-/home/zhuhe/code/litchi/test-data/ole/xls/ConditionalFormattingSamples.xls}
SCRATCH=/tmp/claude-1001/-home-zhuhe-code-litchi/02240c68-6724-42ba-929e-c4157529483d/scratchpad/change-0579
export TMPDIR="$SCRATCH/tmp" RAYON_NUM_THREADS=1 OMP_NUM_THREADS=1
mkdir -p "$TMPDIR"
MODES=${MODES:-file-source owned-readat tracked-file}
for mode in $MODES; do
  for op in open list one-cell; do
    setarch x86_64 -R taskset -c "$CPU" "$ATTR" \
      --input "$INPUT" --mode "$mode" --operation "$op" \
      --warmups "${W:-20}" --samples "${S:-100}" \
      > "$OUT/${mode}-${op}.json" 2> "$OUT/${mode}-${op}.stderr"
    echo "wrote $OUT/${mode}-${op}.json"
  done
done
