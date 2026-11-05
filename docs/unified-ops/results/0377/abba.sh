#!/bin/bash
# Record 0377 wall clock: binary-level A-B-B-A (baseline, candidate,
# candidate, baseline), one pinned CPU; inside every run each route is an
# in-process leaf/unified ABBA of --blocks blocks x --rounds rounds.
F=/deleted-coord-artifact
CPU=${CPU:-47} OUT=${OUT:-$F/abba} BLOCKS=${BLOCKS:-5} ROUNDS=${ROUNDS:-3}
ROUTES=${ROUTES:?route list}
mkdir -p "$OUT"
log() { echo "$* $(date +%T) load $(cut -d' ' -f1-3 /proc/loadavg)" >> "$OUT/log.txt"; }
n=0
for side in base cand cand base; do
  n=$((n+1))
  log "start $n $side"
  taskset -c "$CPU" "$F/bin/fin-$side" --routes "$ROUTES" --blocks "$BLOCKS" --rounds "$ROUNDS" \
    > "$OUT/$side-$n.json" 2> "$OUT/$side-$n.err" || log "FAIL $n $side"
  log "end $n $side"
done
log done
