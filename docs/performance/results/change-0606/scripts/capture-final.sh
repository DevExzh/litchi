#!/usr/bin/env bash
set -u
S=$SCRATCH
R=/home/zhuhe/code/litchi
FIX=$R/test-data/poi/test-data/slideshow/45543.ppt
export TMPDIR=$S/tmp RAYON_NUM_THREADS=1
mkdir -p "$TMPDIR" "$S/cg"
# 1. Callgrind isolation pairs, after leg only (the before binary is unchanged).
B=/home/zhuhe/code/litchi-worktrees/targets/0606-drv-after/release/ppt0606
for mode in eager-open eager-slides eager-text owned-edit-save; do
  for n in 10 110; do
    out=$S/cg/after-$mode-s$n
    setarch x86_64 -R taskset -c 26 valgrind --tool=callgrind \
      --callgrind-out-file=$out.out --cache-sim=no --branch-sim=no \
      "$B" profile "$mode" "$FIX" 1 "$n" > $out.stdout 2> $out.stderr
    echo "rc=$? after-$mode-s$n" >> $S/cg/done2.log
  done
done
# 2. Native instruction/fault counts, both allocator configurations.
: > $S/perf-stat-final.txt
for mode in eager-open eager-slides eager-text owned-edit-save; do
 for leg in before after; do
  BB=/home/zhuhe/code/litchi-worktrees/targets/0606-drv-$leg/release/ppt0606
  for cfg in default heap; do
    if [ $cfg = heap ]; then export MALLOC_MMAP_THRESHOLD_=8388608 MALLOC_TRIM_THRESHOLD_=8388608; else unset MALLOC_MMAP_THRESHOLD_ MALLOC_TRIM_THRESHOLD_; fi
    o=$(taskset -c 26 perf stat -x, -e instructions,cycles,minor-faults -- "$BB" profile $mode "$FIX" 0 1000 2>&1 >/dev/null)
    echo "$mode $leg $cfg loop1000 $(echo "$o" | tr '\n' ' ')" >> $S/perf-stat-final.txt
  done
  unset MALLOC_MMAP_THRESHOLD_ MALLOC_TRIM_THRESHOLD_
  o=$(taskset -c 26 perf stat -x, -r 200 -e instructions,cycles,minor-faults -- "$BB" profile $mode "$FIX" 0 1 2>&1 >/dev/null)
  echo "$mode $leg default single-shot-x200 $(echo "$o" | tr '\n' ' ')" >> $S/perf-stat-final.txt
 done
done
echo FINALCOUNTSDONE >> $S/cg/done2.log
