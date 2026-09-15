#!/usr/bin/env bash
# capture_latency.sh <before-binary> <after-binary> <out-dir>
#
# A1 B1 B2 A2 wall-clock capture. Each round is a separate child; the two
# directions bracket each other, so monotonic drift in the host shows up as a
# disagreement between them, and a1 against a2 (and b1 against b2) is the same
# binary against itself in the same window -- the measured floor.
#
# The probe reports every phase of one open-edit-save separately (open, stage,
# commit, publish), so one run supplies the changed scenario (`commit` of
# `number-source-backed`) and its controls (`open`, and the commit of the two
# publication paths the change does not touch) in the same window.
set -euo pipefail
BEFORE=$(readlink -f "$1")
AFTER=$(readlink -f "$2")
mkdir -p "$3"; OUT=$(readlink -f "$3")
REPO=${REPO:-/home/zhuhe/code/litchi}
CPU=${CPU:-14}
SCRATCH=${SCRATCH:-/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0620}
export TMPDIR="$SCRATCH/tmp" RAYON_NUM_THREADS=1 OMP_NUM_THREADS=1
mkdir -p "$TMPDIR"

# stem:path:warmups:samples
fixtures=(
  "54016:$REPO/test-data/poi/test-data/spreadsheet/54016.xls:5:40"
  "cv:$REPO/test-data/ole/xls/WithCustomViews.xls:20:120"
  "formula:$REPO/test-data/ole/xls/FormulaEvalTestData.xls:20:120"
)

operations_for() {
  case "$1" in
    formula) echo "open number-generic string-generic noop-generic" ;;
    *) echo "open number-plan number-source-backed number-generic string-generic noop-generic" ;;
  esac
}

for round in a1 b1 b2 a2; do
  case "$round" in
    a1|a2) BIN="$BEFORE" ;;
    b1|b2) BIN="$AFTER" ;;
  esac
  for fixture in "${fixtures[@]}"; do
    IFS=: read -r stem path warmups samples <<<"$fixture"
    mkdir -p "$OUT/$round/$stem"
    for op in $(operations_for "$stem"); do
      setarch x86_64 -R taskset -c "$CPU" "$BIN" \
        --input "$path" --operation "$op" \
        --warmups "$warmups" --samples "$samples" \
        > "$OUT/$round/$stem/$op.json" 2>/dev/null
    done
    echo "timed $round/$stem"
  done
done
