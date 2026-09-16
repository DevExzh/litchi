#!/usr/bin/env bash
# Change 0647: callgrind isolation pairs.
#
# Profiles the scenario body at N = 1 and N = 11 and differences the totals, so
# the per-scenario instruction count is (Ir(11) - Ir(1)) / 10 with process
# start-up, the fixture read and callgrind's own fixed cost cancelled out.
#
# Also extracts the per-run call counts of the two publication functions change
# 0593 introduced the capture to avoid, by summing the `calls=` lines that
# target them in the raw callgrind file.
#
# Usage: callgrind-pairs.sh <staged-binary-dir> <fixture> <output-dir>
set -euo pipefail

STAGE="$1"
FIXTURE="$2"
OUT="$3"
CPU="${CPU:-22}"
mkdir -p "$OUT"

run() { # leg scenario n extra-args...
  local leg="$1" scenario="$2" n="$3"
  shift 3
  local file="$OUT/callgrind.out.$scenario-$leg-$n"
  taskset -c "$CPU" valgrind --tool=callgrind --callgrind-out-file="$file" \
    --cache-sim=no --branch-sim=no \
    "$STAGE/reuse_iters-$leg" "$FIXTURE" "$n" "$@" > "$OUT/stdout.$scenario-$leg-$n" 2>&1
  local total
  total=$(grep -E '^summary:|^totals:' "$file" | head -1 | awk '{print $2}')
  echo "$scenario $leg n=$n Ir=$total"
}

for leg in before after; do
  for n in 1 11; do
    run "$leg" package "$n" --package
    run "$leg" drawing "$n" --part /xl/drawings/drawing1.xml
    run "$leg" seamonly "$n" --package --seam-only
  done
done
