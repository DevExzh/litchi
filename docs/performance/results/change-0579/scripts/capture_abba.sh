#!/usr/bin/env bash
# capture_abba.sh <before-binary> <after-binary> <out-dir>
# Paired A/B/B/A medians with a same-binary noise floor measured in the same
# window, the shape change 0576 used. Order per cell:
#   A1 (before) B1 (after) B2 (after) A2 (before)
# dir1 = B1 vs A1, dir2 = B2 vs A2, A/A = A2 vs A1, B/B = B2 vs B1.
set -euo pipefail
BEFORE=$(readlink -f "$1"); AFTER=$(readlink -f "$2"); mkdir -p "$3"; OUT=$(readlink -f "$3")
CPU=${CPU:-17}
SCRATCH=/tmp/claude-1001/-home-zhuhe-code-litchi/02240c68-6724-42ba-929e-c4157529483d/scratchpad/change-0579
export TMPDIR="$SCRATCH/tmp" RAYON_NUM_THREADS=1 OMP_NUM_THREADS=1
mkdir -p "$TMPDIR"
W=${W:-50}; S=${S:-1000}
FLAG=/home/zhuhe/code/litchi/test-data/ole/xls/ConditionalFormattingSamples.xls
CV=/home/zhuhe/code/litchi/test-data/ole/xls/WithCustomViews.xls
K54016=/home/zhuhe/code/litchi/test-data/poi/test-data/spreadsheet/54016.xls
uptime > "$OUT/load-open.txt"
for cell in "flagship:$FLAG" "cv:$CV" "54016:$K54016"; do
  stem=${cell%%:*}; input=${cell#*:}
  for mode in owned-readat file-source; do
    for leg in A1:$BEFORE B1:$AFTER B2:$AFTER A2:$BEFORE; do
      tag=${leg%%:*}; bin=${leg#*:}
      setarch x86_64 -R taskset -c "$CPU" "$bin" \
        --input "$input" --mode "$mode" --operation open \
        --warmups "$W" --samples "$S" \
        > "$OUT/${stem}-${mode}-${tag}.json" 2> "$OUT/${stem}-${mode}-${tag}.stderr"
    done
    echo "wrote $stem/$mode"
  done
done
uptime > "$OUT/load-close.txt"
