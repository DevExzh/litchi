#!/usr/bin/env bash
# Paired timing of the file-source scenarios no selector covers. A1 B1 B2 A2.
set -euo pipefail
S="$1"; FIXTURE="$2"; PART="$3"; SCENARIO="$4"; TAG="$5"; WARMUP="${6:-10}"; SAMPLES="${7:-40}"
B=/home/zhuhe/code/litchi-worktrees/targets/0600-probe-before/release/file_source_time
A=/home/zhuhe/code/litchi-worktrees/targets/0600-probe-after/release/file_source_time
run() { local leg="$1" bin="$2"
  taskset -c 22 "$bin" "$FIXTURE" "$PART" "$SCENARIO" "$WARMUP" "$SAMPLES" \
    > "$S/probe-time-$TAG-$leg.txt" 2>&1
}
run A1 "$B"; run B1 "$A"; run B2 "$A"; run A2 "$B"
