#!/bin/bash
# Record 0362: binary-level A-B-B-A of file_w8_probe, base (trunk f2bbbc8d07) vs head (W8),
# release, pinned to one CPU. Run twice: OUT=run1 and OUT=run2.
W=${W:-/home/zhuhe/code/litchi-worktrees/.coord/w8-file}; CPU=${CPU:-29}; ROUNDS=${ROUNDS:-11}
OUT=${OUT:-$W/abba/run1}; mkdir -p "$OUT"
run() { taskset -c "$CPU" "$W/bin/$1" "$ROUNDS" > "$OUT/$2" 2>>"$OUT/log.txt" || echo "FAIL $2" >> "$OUT/log.txt"; }
echo "start $(date +%T) load $(cut -d' ' -f1 /proc/loadavg)" >> "$OUT/log.txt"
n=0
for leg in before after after before; do
  n=$((n+1)); run $leg probe-$n-$leg.json
  echo "leg $n $leg $(date +%T) load $(cut -d' ' -f1 /proc/loadavg)" >> "$OUT/log.txt"
done
echo done >> "$OUT/log.txt"
