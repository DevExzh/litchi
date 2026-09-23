#!/usr/bin/env bash
# usage: profile_case.sh BIN CASE SHAPE OUTNAME SAMPLES  (frame-pointer perf record + fold)
set -euo pipefail
S=/home/zhuhe/code/litchi-worktrees/scratch/0753
cd $S/prof
taskset -c 8 perf record -q -F 4999 -g -o $4.data "$1" --case "$2" --writer-shape "$3" --samples "$5" --warmup 3 --json $S/tmp/prof-$4.json >/dev/null 2>&1
perf script -i $4.data --inline 2>/dev/null > $4.txt
python3 /home/zhuhe/code/litchi-worktrees/scratch/profile-r2/folded/fold.py $4.txt $4.folded >/dev/null
rm -f $4.txt $4.data
