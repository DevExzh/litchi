#!/usr/bin/env bash
# Paired timing for change 0607. Legs in A1 B1 B2 A2 order; A = resave (the
# authored edit-and-save that materializes the presentation), B = nopsave (the
# same save with a clean model, so flush_presentation returns immediately).
set -euo pipefail
PROBE="$1"; OUT="$2"; SLIDES="$3"; ITERS="$4"; SAMPLES="$5"; CPU="$6"
mkdir -p "$OUT"
run_leg() {
  local name="$1" cmd="$2"
  : > "$OUT/$name.txt"
  for _ in 1 2 3; do taskset -c "$CPU" "$PROBE" "$cmd" "$SLIDES" "$ITERS" >/dev/null; done
  for _ in $(seq 1 "$SAMPLES"); do
    taskset -c "$CPU" "$PROBE" "$cmd" "$SLIDES" "$ITERS" | awk -F'\t' '{print $1/$2}' >> "$OUT/$name.txt"
  done
}
run_leg A1 resave
run_leg B1 nopsave
run_leg B2 nopsave
run_leg A2 resave
