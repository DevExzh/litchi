#!/usr/bin/env bash
# Paired open timing through the retained probe. Order A1 B1 B2 A2.
set -euo pipefail
S="$1"; FIXTURE="$2"; TAG="$3"; SAMPLES="${4:-40}"
B=/home/zhuhe/code/litchi-worktrees/targets/0594-probe-before/release/zip-opc-read-probe
A=/home/zhuhe/code/litchi-worktrees/targets/0594-probe-after/release/zip-opc-read-probe
BW=/home/zhuhe/code/litchi-worktrees/before-08d968f8e
AW=/home/zhuhe/code/litchi-worktrees/0594
run() { local leg="$1" bin="$2" wd="$3"
  ( cd "$wd" && PROBE_PHASES=time PROBE_SAMPLES="$SAMPLES" PROBE_WARMUP=10 \
      taskset -c 14 "$bin" "$FIXTURE" /unused ) > "$S/probe-time-$TAG-$leg.txt" 2>&1
}
run A1 "$B" "$BW"; run B1 "$A" "$AW"; run B2 "$A" "$AW"; run A2 "$B" "$BW"
