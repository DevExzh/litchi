#!/usr/bin/env bash
set -euo pipefail
S="$1"; WARMUP="$2"; SAMPLES="$3"; TAG="$4"
B=/home/zhuhe/code/litchi-worktrees/targets/0594-before/release/litchi-perf-baseline
A=/home/zhuhe/code/litchi-worktrees/targets/0594-after/release/litchi-perf-baseline
BW=/home/zhuhe/code/litchi-worktrees/before-08d968f8e
AW=/home/zhuhe/code/litchi-worktrees/0594
run() { ( cd "$3" && taskset -c 14 "$2" --warmup "$WARMUP" --samples "$SAMPLES" \
    --case xlsx_source_open --xlsx-shape medium --json "$S/abba-$TAG-$1.json" ) > "$S/abba-$TAG-$1.log" 2>&1; }
run A1 "$B" "$BW"; run B1 "$A" "$AW"; run B2 "$A" "$AW"; run A2 "$B" "$BW"
