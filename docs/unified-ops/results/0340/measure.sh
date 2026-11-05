#!/bin/bash
# Record 0340 measurement driver (record 0337's set A for the iWork formats
# and ODC): three trials (processes) per binary, one pinned CPU, each case an
# in-process ABBA of 5 rounds (10 samples a leg) after an untimed
# equivalence check; fixture of 10,000 targets, 100 / 1,000 / 10,000
# operations. The before and after binaries alternate trial by trial so both
# see the same host load.
set -u
BEFORE=$1 AFTER=$2 OUT=$3 CPU=${CPU:-30}
mkdir -p "$OUT"
log() { echo "== $* $(date +%T) load $(cut -d' ' -f1 /proc/loadavg)" >> "$OUT/log.txt"; }
run() { local name=$1 bin=$2; shift 2; log "$name $*"; taskset -c "$CPU" "$bin" "$@" >> "$OUT/$name" 2>> "$OUT/log.txt" || echo "FAIL $name" >> "$OUT/log.txt"; }
F=pages,keynote,numbers,odc
for trial in 1 2 3; do
  run batch-A-before-$trial.json "$BEFORE" --formats $F --counts 100,1000,10000 --rounds 5
  run batch-A-after-$trial.json "$AFTER" --formats $F --counts 100,1000,10000 --rounds 5
done
echo done >> "$OUT/log.txt"
