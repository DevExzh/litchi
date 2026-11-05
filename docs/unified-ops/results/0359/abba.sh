#!/bin/bash
# Record 0359: binary-level A-B-B-A of legacy_w5_probe, base (trunk) vs head (W5),
# release, pinned to one CPU. Binaries: perf/tgt-{before,after}/release/examples/legacy_w5_probe
# (source: crates/litchi-ops/examples/legacy_w5_probe.rs on both trees).
P=${P:-/home/zhuhe/code/litchi-worktrees/.coord/w5-legacy/perf}; CPU=${CPU:-29}; ROUNDS=${ROUNDS:-31}
OUT=${OUT:-$P/out}; mkdir -p "$OUT"
run() { taskset -c "$CPU" "$P/tgt-$1/release/examples/legacy_w5_probe" "$ROUNDS" > "$OUT/$2" 2>>"$OUT/log.txt" || echo "FAIL $2" >> "$OUT/log.txt"; }
echo "start $(date +%T) load $(cut -d' ' -f1 /proc/loadavg)" >> "$OUT/log.txt"
n=0
for leg in before after after before; do
  n=$((n+1)); run $leg probe-$leg-$n.json
  echo "leg $n $leg $(date +%T) load $(cut -d' ' -f1 /proc/loadavg)" >> "$OUT/log.txt"
done
echo done >> "$OUT/log.txt"
