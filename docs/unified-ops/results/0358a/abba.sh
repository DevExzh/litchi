#!/bin/bash
# Record 0358a: binary-level A-B-B-A, release, pinned to one CPU.
#  - odf_source_backed_probe: before = base e4a0614f5d with probe_before.rs as
#    examples/odf_source_backed_probe.rs (the former stream routes); after =
#    the 0358a head (crates/litchi-ops/examples/odf_source_backed_probe.rs).
#  - odf_w6_probe: before = record 0358 head f755903f45; after = the 0358a head.
# Binaries: $P/tgt-{before,w6base,after}/release/examples/*; hashes in binaries.sha256.
# Each probe dumps its first-round outputs once per leg (sb-<leg>-<n>/) for the
# byte-identity check of analyze.py.
P=${P:-/home/zhuhe/code/litchi-worktrees/.coord/w6-odf/perf}; CPU=${CPU:-29}
SB_ROUNDS=${SB_ROUNDS:-15}; W6_ROUNDS=${W6_ROUNDS:-61}
OUT=${OUT:-$P/out-0358a}; mkdir -p "$OUT"
cd /home/zhuhe/code/litchi-worktrees/uo2-w6-odf/crates/litchi-ops
log() { echo "$* $(date +%T) load $(cut -d' ' -f1 /proc/loadavg)" >> "$OUT/log.txt"; }
log start
n=0
for leg in before after after before; do
  n=$((n+1))
  w6=$([ "$leg" = before ] && echo w6base || echo after)
  taskset -c "$CPU" "$P/tgt-$leg/release/examples/odf_source_backed_probe" "$SB_ROUNDS" "$OUT/sb-$leg-$n" \
    > "$OUT/sb-$leg-$n.json" 2>>"$OUT/log.txt" || log "FAIL sb $leg $n"
  taskset -c "$CPU" "$P/tgt-$w6/release/examples/odf_w6_probe" "$W6_ROUNDS" \
    > "$OUT/w6-$leg-$n.json" 2>>"$OUT/log.txt" || log "FAIL w6 $leg $n"
  log "leg $n $leg"
done
log done
