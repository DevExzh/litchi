#!/usr/bin/env bash
# Change 0744 confirmation ABBA for the flagged no-op pairs and the eager
# control's after-arm spread (same binaries, CPU 12, A1 B1 B2 A2 A3 B3 B4 A4).
set -euo pipefail
A=${A:?}; B=${B:?}; OUT=${OUT:-confirm}
mkdir -p "$OUT"
for slot in A1 B1 B2 A2 A3 B3 B4 A4; do
  bin=$A; [ "${slot:0:1}" = B ] && bin=$B
  taskset -c 12 "$bin" --case xlsx_noop_commit_save --xlsx-shape dense-wide,medium --samples 2000 --warmup 200 --json "$OUT/noop-$slot.json" > /dev/null 2> "$OUT/noop-$slot.stderr"
  taskset -c 12 "$bin" --case xlsx_eager_cell_values_one_edit_save --xlsx-cell-crud-shape medium --samples 100 --warmup 10 --json "$OUT/eager-$slot.json" > /dev/null 2> "$OUT/eager-$slot.stderr"
  echo "$(date +%T) confirm $slot done"
done
