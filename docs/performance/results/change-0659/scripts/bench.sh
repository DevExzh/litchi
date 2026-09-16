#!/usr/bin/env bash
# Paired timing, order A1 B1 B2 A2, pinned to CPU 14.
#   bench.sh <before-binary> <after-binary> <mode> <fixture> <stem> <outdir> <warmups> <samples>
set -euo pipefail
BEFORE=$1; AFTER=$2; MODE=$3; FIXTURE=$4; STEM=$5; OUT=$6; W=$7; N=$8
mkdir -p "$OUT"
for leg in A1:$BEFORE B1:$AFTER B2:$AFTER A2:$BEFORE; do
  taskset -c 14 "${leg#*:}" bench "$MODE" "$FIXTURE" "$W" "$N" \
    > "$OUT/bench-$MODE-$STEM-${leg%%:*}.txt" 2>/dev/null
done
