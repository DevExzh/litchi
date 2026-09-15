#!/usr/bin/env bash
# Native `perf stat` isolation pairs for the length-changing OLE2 edit-and-save.
#
# Callgrind prices `rep movsb`/`rep stosb` once per byte, and the container leg
# is 85% bulk copy, so its callgrind share is an upper bound. This script prices
# the same legs in retired cycles and instructions.
#
# Each case is run at LOW and HIGH iterations, REPS times each; the medians are
# differenced and divided by (HIGH - LOW), so process start-up, the fixture read
# and the probe's own printing cancel.
set -euo pipefail

PROBE=${PROBE:?}
REPO=${REPO:?}
OUT=${OUT:?}
AUTHORED=${AUTHORED:?}
CPU=${CPU:-11}
LOW=${LOW:-10}
HIGH=${HIGH:-110}
REPS=${REPS:-11}

mkdir -p "$OUT/cycles"

measure() { # name fmt input op n
  local name=$1 fmt=$2 input=$3 op=$4 n=$5
  shift 5
  local raw="$OUT/cycles/${name}-${op}-s${n}.csv"
  : >"$raw"
  local i
  for ((i = 0; i < REPS; i++)); do
    taskset -c "$CPU" perf stat -e cycles,instructions -x, \
      "$PROBE" --format "$fmt" --input "$input" --operation "$op" \
      --warmups 0 --samples "$n" "$@" 2>>"$raw" >/dev/null
  done
}

for spec in \
  "xls54016|xls|test-data/poi/test-data/spreadsheet/54016.xls|" \
  "docfloat|doc|test-data/ole/doc/FloatingPictures.doc|" \
  "docnohf|doc|test-data/ole/doc/NoHeadFoot.doc|" \
  "ppt45543|ppt|test-data/poi/test-data/slideshow/45543.ppt|--ppt-op remove-slide --slide 1" \
  "pptauthored|ppt|__AUTHORED__|--ppt-op text --slide 1 --shape 0"
do
  IFS='|' read -r name fmt fixture extra <<<"$spec"
  input="$REPO/$fixture"
  if [ "$fixture" = "__AUTHORED__" ]; then input="$AUTHORED"; fi
  for op in open commit container container-changed-only; do
    # shellcheck disable=SC2086
    measure "$name" "$fmt" "$input" "$op" "$LOW" $extra
    # shellcheck disable=SC2086
    measure "$name" "$fmt" "$input" "$op" "$HIGH" $extra
    echo "done $name $op" >&2
  done
done

# An A/A control on the same metric: the same leg measured twice, interleaved
# with the runs above, so that drift on this shared host is visible.
for i in A B; do
  measure "aafloor$i" doc "$REPO/test-data/ole/doc/FloatingPictures.doc" commit "$LOW"
  measure "aafloor$i" doc "$REPO/test-data/ole/doc/FloatingPictures.doc" commit "$HIGH"
done
