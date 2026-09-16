#!/usr/bin/env bash
# Change 0647: paired timing, ordered A1 B1 B2 A2 (before, after, after,
# before) so a monotone drift in host load cancels, plus an A/A floor measured
# in the same window by running the before binary against itself in the same
# four-block order.
#
# Both binaries are staged outside any Cargo target directory (change 0627: a
# concurrent build relinked one mid-run) and pinned to one CPU.
#
# Usage: timing.sh <staged-binary-dir> <fixture> <output-dir> <samples> [args...]
set -euo pipefail

STAGE="$1"
FIXTURE="$2"
OUT="$3"
SAMPLES="$4"
shift 4
CPU="${CPU:-22}"
WARMUP="${WARMUP:-20}"
mkdir -p "$OUT"

run_block() { # tag leg
  local tag="$1" leg="$2"
  taskset -c "$CPU" "$STAGE/reuse_iters-$leg" "$FIXTURE" 1 \
    --timing "$SAMPLES" --warmup "$WARMUP" "${EXTRA[@]}" 2>/dev/null > "$OUT/$tag.txt"
}

EXTRA=("$@")

run_block a1 before
run_block b1 after
run_block b2 after
run_block a2 before

cat "$OUT/a1.txt" "$OUT/a2.txt" > "$OUT/before.txt"
cat "$OUT/b1.txt" "$OUT/b2.txt" > "$OUT/after.txt"

# A/A floor in the same window: the same four-block order, before against
# before, so the reported floor is this window's own noise.
run_block aa1 before
run_block aa2 before
run_block aa3 before
run_block aa4 before
cat "$OUT/aa1.txt" "$OUT/aa4.txt" > "$OUT/floor-a.txt"
cat "$OUT/aa2.txt" "$OUT/aa3.txt" > "$OUT/floor-b.txt"
