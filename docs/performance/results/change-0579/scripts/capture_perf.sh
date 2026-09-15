#!/usr/bin/env bash
# capture_perf.sh <attribution-binary> <out-dir> [tag]
# perf stat isolation pairs: differencing an 1100-sample and a 100-sample child
# and dividing by 1000 isolates one operation. Change 0564's method, reused
# unchanged by changes 0574 and 0576.
set -euo pipefail
ATTR=$(readlink -f "$1"); mkdir -p "$2"; OUT=$(readlink -f "$2"); TAG=${3:-leg}
CPU=${CPU:-17}
SCRATCH=/tmp/claude-1001/-home-zhuhe-code-litchi/02240c68-6724-42ba-929e-c4157529483d/scratchpad/change-0579
export TMPDIR="$SCRATCH/tmp" RAYON_NUM_THREADS=1 OMP_NUM_THREADS=1
mkdir -p "$TMPDIR"
EVENTS=instructions,cycles,branches,branch-misses,page-faults,task-clock
MODES=${MODES:-owned-readat file-source}
OPS=${OPS:-open}
INPUTS=${INPUTS:-/home/zhuhe/code/litchi/test-data/ole/xls/ConditionalFormattingSamples.xls}
for input in $INPUTS; do
  stem=$(basename "$input" .xls)
  for mode in $MODES; do
    for op in $OPS; do
      for samples in 100 1100; do
        setarch x86_64 -R taskset -c "$CPU" \
          perf stat -x, -e "$EVENTS" -o "$OUT/${TAG}-${stem}-${mode}-${op}-s${samples}.csv" -- \
          "$ATTR" --input "$input" --mode "$mode" --operation "$op" \
          --warmups 5 --samples "$samples" > /dev/null 2>"$OUT/${TAG}-${stem}-${mode}-${op}-s${samples}.stderr"
      done
      echo "wrote $OUT/${TAG}-${stem}-${mode}-${op}-s{100,1100}.csv"
    done
  done
done
