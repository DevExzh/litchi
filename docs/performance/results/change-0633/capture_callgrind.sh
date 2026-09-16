#!/usr/bin/env bash
# capture_callgrind.sh <probe-binary> <leg-name> <out-dir>
#
# Per-phase instruction attribution of the XLS edit-and-save path, by change
# 0574's isolation method as 0576, 0584, 0595 and 0605 used it: run the same
# child at a small and a large sample count, difference the two profiles and
# divide by the extra operations. Warmups and the probe's own target-discovery
# open are identical in both legs, so everything outside the measured operation
# cancels.
#
# `--separate-callers=2` is what makes this an *attribution*: it splits each
# callee by its two-deep caller chain, so the three complete `Workbook::new`
# parses a source-backed commit runs are reported as three separate lines with
# their own call sites rather than summed under one symbol.
#
# Callgrind runs SHA-256 in software (valgrind masks the SHA CPUID bit), so the
# fingerprint share here is an upper bound; capture_perf.sh takes the native
# cycle counts beside it.
set -euo pipefail
BIN=$(readlink -f "$1")
LEG="$2"
mkdir -p "$3"; OUT=$(readlink -f "$3")
REPO=${REPO:-/home/zhuhe/code/litchi}
CPU=${CPU:-9}
SCRATCH=${SCRATCH:-/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0633}
export TMPDIR="$SCRATCH/tmp" RAYON_NUM_THREADS=1 OMP_NUM_THREADS=1
mkdir -p "$TMPDIR"

# stem:path:warmups:small:large
fixtures=(
  "54016:$REPO/test-data/poi/test-data/spreadsheet/54016.xls:1:1:4"
  "cv:$REPO/test-data/ole/xls/WithCustomViews.xls:2:3:12"
  "formula:$REPO/test-data/ole/xls/FormulaEvalTestData.xls:2:3:12"
)

# FormulaEvalTestData.xls carries a VBA project, so both source-backed
# publication paths refuse it before planning; only the generic paths run.
operations_for() {
  case "$1" in
    formula) echo "open number-generic string-generic noop-generic" ;;
    *) echo "open number-plan number-source-backed number-generic string-generic noop-generic" ;;
  esac
}

run_pair() {
  local stem=$1 path=$2 warmups=$3 op=$4 small=$5 large=$6
  for pair in "small:$small" "large:$large"; do
    local size=${pair%%:*} samples=${pair#*:}
    local raw="$OUT/cg-$LEG-$stem-$op-$size.out"
    setarch x86_64 -R taskset -c "$CPU" \
      valgrind --tool=callgrind --callgrind-out-file="$raw" --quiet \
      --cache-sim=no --branch-sim=no --separate-callers=2 \
      "$BIN" --input "$path" --operation "$op" \
      --warmups "$warmups" --samples "$samples" \
      > /dev/null 2> "$OUT/cg-$LEG-$stem-$op-$size.stderr"
    # Keep only the function-total section: the per-source auto-annotation
    # that callgrind_annotate appends is 1.8 MB per file and nothing reads it.
    callgrind_annotate --inclusive=yes --threshold=99.9 "$raw" \
      | sed '/^-- Auto-annotated source/,$d' \
      > "$OUT/ann-$LEG-$stem-$op-$size.txt"
    rm -f "$raw"
    echo "annotated $LEG/$stem/$op/$size ($samples samples)"
  done
  echo "$stem $op $small $large" >> "$OUT/pairs-$LEG.txt"
}

: > "$OUT/pairs-$LEG.txt"
for fixture in "${fixtures[@]}"; do
  IFS=: read -r stem path warmups small large <<<"$fixture"
  for op in $(operations_for "$stem"); do
    run_pair "$stem" "$path" "$warmups" "$op" "$small" "$large"
  done
done
