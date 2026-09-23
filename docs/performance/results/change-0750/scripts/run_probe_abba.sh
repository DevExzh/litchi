#!/bin/bash
# Change 0750 audit probe: the same probe source built against each leg's
# xml-minifier; four rounds of A B B A processes, pinned to one core. Each
# process times every case in 41 batches of about 20 ms.
set -u
BEFORE=${BEFORE:?}
AFTER=${AFTER:?}
PARTS=${PARTS:?}
OUT=${OUT:?}
CORE=${CORE:-24}
mkdir -p "$OUT/raw"
for round in 1 2 3 4; do
  for slot in 1 2 3 4; do
    case $slot in 1|4) arm=A; bin=$BEFORE;; *) arm=B; bin=$AFTER;; esac
    taskset -c "$CORE" "$bin" "$PARTS" > "$OUT/raw/probe-r$round-s$slot-$arm.jsonl" 2> "$OUT/raw/probe-r$round-s$slot-$arm.log"
    echo "round=$round slot=$slot arm=$arm exit=$?" >> "$OUT/status.txt"
  done
done
echo done >> "$OUT/status.txt"
