#!/bin/bash
# Record 0360b: binary-level A-B-B-A, base (before) vs head (after), release,
# pinned to one CPU. Binaries: perf/tgt-{before,after}/release/...
P=${P:-/deleted-coord-artifact}; CPU=${CPU:-29}; ROUNDS=${ROUNDS:-31}
OUT=${OUT:-$P/out}; mkdir -p "$OUT"
FIXTURE=${FIXTURE:-/home/zhuhe/code/litchi-worktrees/uo2-w4b-xlsb/test-data/poi/test-data/spreadsheet/testVarious.xlsb}
probe() { taskset -c "$CPU" "$P/tgt-$1/release/examples/xlsb_w4b_probe" "$ROUNDS" > "$OUT/probe-$1-$2.json" 2>>"$OUT/log.txt" || echo "FAIL probe $1 $2" >> "$OUT/log.txt"; }
sheet() { taskset -c "$CPU" "$P/tgt-$1/release/examples/sheet_dispatch_bench" > "$OUT/sheet-$1-$2.txt" 2>>"$OUT/log.txt" || echo "FAIL sheet $1 $2" >> "$OUT/log.txt"; }
crud() { taskset -c "$CPU" "$P/tgt-$1/release/xlsb_crud" --case all --backend owned --fixture "$FIXTURE" --warmup 3 --samples "$ROUNDS" --json "$OUT/crud-$1-$2.json" > /dev/null 2>>"$OUT/log.txt" || echo "FAIL crud $1 $2" >> "$OUT/log.txt"; }
echo "start $(date +%T) load $(cut -d' ' -f1 /proc/loadavg)" >> "$OUT/log.txt"
n=0
for leg in before after after before; do
  n=$((n+1)); probe $leg $n; sheet $leg $n; crud $leg $n
  echo "leg $n $leg $(date +%T) load $(cut -d' ' -f1 /proc/loadavg)" >> "$OUT/log.txt"
done
echo done >> "$OUT/log.txt"
