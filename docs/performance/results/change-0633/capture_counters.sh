#!/usr/bin/env bash
# capture_counters.sh <probe-binary> <leg-name> <out-dir>
#
# Deterministic counters first: allocation calls, allocated bytes and peak live
# bytes for one open and one whole open-edit-save, the published byte count, and
# the source-backed overlay diagnostics (splices, replacement bytes, changed
# spans, source and target Workbook lengths). The probe's counting allocator is
# armed only for the first measured iteration, so every figure is per operation
# and excludes the process's own start-up and the fixture copy.
#
# `inventory` runs first and records the per-worksheet record census that names
# each fixture's cell mix and the two targets the edits use.
set -euo pipefail
BIN=$(readlink -f "$1")
LEG="$2"
mkdir -p "$3"; OUT=$(readlink -f "$3")
REPO=${REPO:-/home/zhuhe/code/litchi}
CPU=${CPU:-9}
SCRATCH=${SCRATCH:-/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0633}
export TMPDIR="$SCRATCH/tmp" RAYON_NUM_THREADS=1 OMP_NUM_THREADS=1
mkdir -p "$TMPDIR"

fixtures=(
  "54016:$REPO/test-data/poi/test-data/spreadsheet/54016.xls"
  "cv:$REPO/test-data/ole/xls/WithCustomViews.xls"
  "formula:$REPO/test-data/ole/xls/FormulaEvalTestData.xls"
  "flagship:$REPO/test-data/ole/xls/ConditionalFormattingSamples.xls"
  "59858:$REPO/test-data/poi/test-data/spreadsheet/59858.xls"
)

: > "$OUT/counters-$LEG.jsonl"
: > "$OUT/inventory-$LEG.jsonl"
for fixture in "${fixtures[@]}"; do
  IFS=: read -r stem path <<<"$fixture"
  if inventory=$(setarch x86_64 -R taskset -c "$CPU" "$BIN" --input "$path" --operation inventory 2>&1); then
    echo "$inventory" >> "$OUT/inventory-$LEG.jsonl"
  else
    echo "{\"stem\":\"$stem\",\"refused\":\"$(printf "%s" "$inventory" | tr '"' "'" | tr -d "\n")\"}" >> "$OUT/inventory-$LEG.jsonl"
    echo "refused $stem: $inventory"
    continue
  fi
  for op in open number-plan number-source-backed number-generic string-generic noop-generic; do
    if line=$(setarch x86_64 -R taskset -c "$CPU" "$BIN" --input "$path" \
        --operation "$op" --warmups 2 --samples 5 2>&1); then
      echo "$line" >> "$OUT/counters-$LEG.jsonl"
    else
      echo "{\"stem\":\"$stem\",\"operation\":\"$op\",\"refused\":\"$(printf "%s" "$line" | tr '"' "'" | tr -d "\n")\"}" \
        >> "$OUT/counters-$LEG.jsonl"
    fi
    echo "counted $LEG/$stem/$op"
  done
done
