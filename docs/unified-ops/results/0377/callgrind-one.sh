#!/bin/bash
# cg-one.sh <side base|cand> <group> <route> <leg>: one callgrind run of one
# route leg; counts only fin_perf_region (the timed calls).
set -u
F=/home/zhuhe/code/litchi-worktrees/.coord/fin-perf
side=$1 group=$2 route=$3 leg=$4
format=${route%%_*}
iters=3; [[ $group == batch ]] && iters=1
out=$F/${CGDIR:-cg}/$side/$route.$leg
mkdir -p $F/${CGDIR:-cg}/$side
systemd-run --user --scope --quiet --collect -p OOMPolicy=continue -p MemoryMax=20G -p MemorySwapMax=0 -- \
  valgrind --tool=callgrind --collect-atstart=no --toggle-collect='*fin_perf_region*' \
  --callgrind-out-file=$out.cg $F/bin/fin-$side --groups $group --formats $format --routes $route \
  --callgrind $leg --iterations $iters > $out.log 2>&1
rc=$?
total=$(grep -m1 '^totals:' $out.cg 2>/dev/null | awk '{print $2}')
echo -e "$side\t$group\t$route\t$leg\t$iters\t${total:-NA}\t$rc"
