#!/usr/bin/env bash
# Callgrind isolation pairs for the length-changing OLE2 edit-and-save probe.
#
# Each case is profiled at LOW and HIGH iterations; the per-symbol self costs
# are differenced and divided by (HIGH - LOW), so process start-up, fixture
# reads and the probe's own printing cancel out.
set -euo pipefail

PROBE=${PROBE:?}
REPO=${REPO:?}
OUT=${OUT:?}
CPU=${CPU:-11}
AUTHORED=${AUTHORED:?}
LOW=${LOW:-2}
HIGH=${HIGH:-12}

mkdir -p "$OUT/raw"

run() {
  local name=$1 fmt=$2 fixture=$3 op=$4
  shift 4
  for n in "$LOW" "$HIGH"; do
    local file="$OUT/raw/${name}-${op}-s${n}.out"
    if [ -s "$file" ]; then continue; fi
    local input="$REPO/$fixture"
    if [ "$fixture" = "__AUTHORED__" ]; then input="$AUTHORED"; fi
    taskset -c "$CPU" valgrind --tool=callgrind --callgrind-out-file="$file" \
      --cache-sim=no --branch-sim=no \
      "$PROBE" --format "$fmt" --input "$input" --operation "$op" \
      --warmups 0 --samples "$n" "$@" >/dev/null 2>"$OUT/raw/${name}-${op}-s${n}.log"
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
  for op in open commit container container-changed-only; do
    # shellcheck disable=SC2086
    run "$name" "$fmt" "$fixture" "$op" $extra
    echo "done $name $op" >&2
  done
done
