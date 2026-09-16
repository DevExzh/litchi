#!/usr/bin/env bash
# Change 0642 paired-timing driver.
#
# usage: abba.sh TAG OPERATION FILE SHEET WARMUP SAMPLES A_BINARY B_BINARY [CPU]
#
# Runs the staged probe binaries in A1 B1 B2 A2 order, one process per leg, each
# pinned to the same CPU. Every leg writes one elapsed nanosecond count per line
# into runs/TAG-{a1,b1,b2,a2}.txt. Passing the same binary as A and B gives the
# A/A floor for that scenario in the same window.
set -euo pipefail
STAGE="$(cd "$(dirname "$0")" && pwd)"
RUNS="$STAGE/runs"
mkdir -p "$RUNS"
TAG="$1"; OPERATION="$2"; FILE="$3"; SHEET="$4"; WARMUP="$5"; SAMPLES="$6"
A="$7"; B="$8"; CPU="${9:-17}"
leg() {
  taskset -c "$CPU" "$STAGE/$2" bench "$OPERATION" "$FILE" "$SHEET" "$WARMUP" "$SAMPLES" \
    > "$RUNS/$TAG-$1.txt" 2> "$RUNS/$TAG-$1.err"
}
leg a1 "$A"; leg b1 "$B"; leg b2 "$B"; leg a2 "$A"
echo "$TAG done"
