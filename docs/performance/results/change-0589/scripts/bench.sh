#!/usr/bin/env bash
# Paired timing, order A1 B1 B2 A2, plus an A/A leg in the same window.
SC=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0589
BEF=/home/zhuhe/code/litchi-worktrees/targets/0589-before/release/snapfence_probe
AFT=/home/zhuhe/code/litchi-worktrees/0589/target/release/snapfence_probe
mode=$1; path=$2; warm=$3; n=$4; stem=$5
run() { taskset -c 9 "$1" bench "$mode" "$path" "$warm" "$n" 2> /dev/null; }
run $BEF > $SC/out/bench-$stem-A1.txt
run $AFT > $SC/out/bench-$stem-B1.txt
run $AFT > $SC/out/bench-$stem-B2.txt
run $BEF > $SC/out/bench-$stem-A2.txt
run $BEF > $SC/out/bench-$stem-A3.txt
run $BEF > $SC/out/bench-$stem-A4.txt
