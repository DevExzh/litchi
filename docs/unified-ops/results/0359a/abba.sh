#!/bin/bash
# Record 0359a: binary-level A-B-B-A of legacy_w5_source_probe and
# legacy_w5_probe, base (7afba5f146, record 0359 tip) vs head (0359a),
# release, pinned to one CPU, each run in a 20 GB memory scope.
P=${P:-/home/zhuhe/code/litchi-worktrees/.coord/w5-legacy/perf2}; CPU=${CPU:-29}; ROUNDS=${ROUNDS:-31}
OUT=${OUT:-$P/out}; mkdir -p "$OUT"
export LITCHI_PROBE_BATCH_ROUNDS=${LITCHI_PROBE_BATCH_ROUNDS:-7}
run() {
  systemd-run --user --scope --quiet --collect -p OOMPolicy=continue -p MemoryMax=20G -p MemorySwapMax=0 -- \
    /usr/bin/time -f "$3 $2 maxrss_kb=%M wall=%e" taskset -c "$CPU" "$P/tgt-$1/release/examples/$3" "$ROUNDS" \
    > "$OUT/$2" 2>>"$OUT/log.txt" || echo "FAIL $2" >> "$OUT/log.txt"
}
echo "start $(date +%T) load $(cut -d' ' -f1 /proc/loadavg)" >> "$OUT/log.txt"
n=0
for leg in before after after before; do
  n=$((n+1))
  run $leg source-$leg-$n.json legacy_w5_source_probe
  run $leg ops-$leg-$n.json legacy_w5_probe
  echo "leg $n $leg $(date +%T) load $(cut -d' ' -f1 /proc/loadavg)" >> "$OUT/log.txt"
done
echo done >> "$OUT/log.txt"
