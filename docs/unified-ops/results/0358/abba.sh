#!/bin/bash
# Record 0358: binary-level A-B-B-A of odf_w6_probe, base (trunk) vs head (W6),
# release, pinned to one CPU. Binaries: perf/tgt-{before,after}/release/examples/odf_w6_probe
# (source: crates/litchi-ops/examples/odf_w6_probe.rs on both trees).
P=${P:-/deleted-coord-artifact}; CPU=${CPU:-29}; ROUNDS=${ROUNDS:-15}
OUT=${OUT:-$P/out}; mkdir -p "$OUT"
cd /home/zhuhe/code/litchi-worktrees/uo2-w6-odf/crates/litchi-ops
run() { taskset -c "$CPU" "$P/tgt-$1/release/examples/odf_w6_probe" "$ROUNDS" > "$OUT/$2" 2>>"$OUT/log.txt" || echo "FAIL $2" >> "$OUT/log.txt"; }
echo "start $(date +%T) load $(cut -d' ' -f1 /proc/loadavg)" >> "$OUT/log.txt"
n=0
for leg in before after after before; do
  n=$((n+1)); run $leg probe-$leg-$n.json
  echo "leg $n $leg $(date +%T) load $(cut -d' ' -f1 /proc/loadavg)" >> "$OUT/log.txt"
done
echo done >> "$OUT/log.txt"
