#!/usr/bin/env bash
# capture_counters.sh <attribution-binary> <leg-name> <out-dir>
#
# Deterministic logical counters (reads, read bytes, version observations, len
# observations, seeks) for the three fixtures change 0605 measures, over the two
# in-memory source modes and the operations the leg's binary supports. These are
# the oracle for "not one byte of I/O moved on the existing paths".
#
# `open`, `list` and `one-cell` exist in both legs. `second-cell`, `full-text`
# and `all-cells` exist only in the after leg's binary; the loop skips an
# operation the binary rejects, so the same script drives both legs.
#
# `54016.xls` has one worksheet, so its per-sheet operations are driven with
# `--worksheet-index 0`; the other two fixtures keep the harness default of 1.
set -euo pipefail
ATTR=$(readlink -f "$1")
LEG="$2"
mkdir -p "$3"; OUT=$(readlink -f "$3")
REPO=${REPO:-/home/zhuhe/code/litchi}
CPU=${CPU:-25}
SCRATCH=${SCRATCH:-/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0605}
export TMPDIR="$SCRATCH/tmp" RAYON_NUM_THREADS=1 OMP_NUM_THREADS=1
mkdir -p "$TMPDIR"

# stem:path:worksheet-index
cells=(
  "flagship:$REPO/test-data/ole/xls/ConditionalFormattingSamples.xls:1"
  "cv:$REPO/test-data/ole/xls/WithCustomViews.xls:1"
  "54016:$REPO/test-data/poi/test-data/spreadsheet/54016.xls:0"
)

for cell in "${cells[@]}"; do
  IFS=: read -r stem path sheet <<<"$cell"
  for mode in owned-readat file-source; do
    for op in open list one-cell second-cell full-text all-cells all-cells-per-cell; do
      extra=()
      name="$op"
      if [ "$op" = all-cells-per-cell ]; then
        op=all-cells; extra=(--all-cells-strategy per-cell)
      elif [ "$op" = all-cells ]; then
        extra=(--all-cells-strategy scan)
      fi
      if setarch x86_64 -R taskset -c "$CPU" "$ATTR" \
        --input "$path" --mode "$mode" --operation "$op" "${extra[@]}" \
        --worksheet-index "$sheet" --warmups "${W:-20}" --samples "${S:-100}" \
        > "$OUT/$LEG-$stem-$mode-$name.json" 2> "$OUT/$LEG-$stem-$mode-$name.stderr"
      then
        echo "wrote $OUT/$LEG-$stem-$mode-$name.json"
      else
        rm -f "$OUT/$LEG-$stem-$mode-$name.json"
        echo "skipped $LEG/$stem/$mode/$name ($(head -c 120 "$OUT/$LEG-$stem-$mode-$name.stderr" | tr '\n' ' '))"
      fi
    done
  done
done
