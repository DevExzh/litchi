#!/bin/bash
F=/home/zhuhe/code/litchi-worktrees/.coord/fin-perf
side=$1 case=$2 n=$3
out=$F/cgpb/$side.$case.$n
mkdir -p $F/cgpb
systemd-run --user --scope --quiet --collect -p MemoryMax=20G -p MemorySwapMax=0 -- valgrind --tool=callgrind --callgrind-out-file=$out.cg $F/bin/pb-$side --case $case --samples $n --warmup 1 --json /dev/null > $out.log 2>&1
echo -e "$side\t$case\t$n\t$(grep -m1 '^totals:' $out.cg | awk '{print $2}')\t$?"
