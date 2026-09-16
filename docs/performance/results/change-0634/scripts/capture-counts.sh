#!/usr/bin/env bash
# Deterministic counts for change 0634: callgrind isolation pairs (s=10 vs s=110)
# and native perf-stat counts, for one leg.  usage: capture-counts.sh <leg> <cpu>
set -u
LEG=$1
CPU=$2
S=${SCRATCH:?}
R=/home/zhuhe/code/litchi
FIX=$R/test-data/poi/test-data/slideshow/45543.ppt
NFIX=$R/test-data/poi/test-data/slideshow/headers_footers_2007.ppt
BIN=$S/bin/ppt0634-$LEG
export TMPDIR=$S/tmp RAYON_NUM_THREADS=1
mkdir -p "$TMPDIR" "$S/cg" "$S/out"
for spec in eager-open:$FIX eager-slides:$FIX eager-text:$FIX eager-notes:$NFIX owned-edit-save:$FIX; do
  mode=${spec%%:*}; fx=${spec#*:}
  for n in 10 110; do
    out=$S/cg/$LEG-$mode-s$n
    setarch x86_64 -R taskset -c $CPU valgrind --tool=callgrind \
      --callgrind-out-file=$out.out --cache-sim=no --branch-sim=no \
      "$BIN" profile "$mode" "$fx" 1 "$n" > $out.stdout 2> $out.stderr
    echo "rc=$? $LEG-$mode-s$n" >> $S/cg/done.log
  done
done
: > $S/out/perf-stat-$LEG.txt
for spec in eager-open:$FIX eager-slides:$FIX eager-text:$FIX eager-notes:$NFIX owned-edit-save:$FIX; do
  mode=${spec%%:*}; fx=${spec#*:}
  for cfg in default heap; do
    if [ $cfg = heap ]; then export MALLOC_MMAP_THRESHOLD_=8388608 MALLOC_TRIM_THRESHOLD_=8388608
    else unset MALLOC_MMAP_THRESHOLD_ MALLOC_TRIM_THRESHOLD_; fi
    o=$(taskset -c $CPU perf stat -x, -e instructions,cycles,minor-faults -- "$BIN" profile $mode "$fx" 0 1000 2>&1 >/dev/null)
    echo "$mode $LEG $cfg loop1000 $(echo "$o" | tr '\n' ' ')" >> $S/out/perf-stat-$LEG.txt
  done
  unset MALLOC_MMAP_THRESHOLD_ MALLOC_TRIM_THRESHOLD_
  o=$(taskset -c $CPU perf stat -x, -r 200 -e instructions,cycles,minor-faults -- "$BIN" profile $mode "$fx" 0 1 2>&1 >/dev/null)
  echo "$mode $LEG default single-shot-x200 $(echo "$o" | tr '\n' ' ')" >> $S/out/perf-stat-$LEG.txt
done
# Allocation counts.
: > $S/out/allocations-$LEG.txt
for spec in eager-open:$FIX eager-slides:$FIX eager-text:$FIX eager-notes:$NFIX owned-edit-save:$FIX source-open:$FIX; do
  mode=${spec%%:*}; fx=${spec#*:}
  taskset -c $CPU "$S/bin/ppt0634_alloc-$LEG" "$mode" "$fx" >> $S/out/allocations-$LEG.txt 2>&1
done
echo "COUNTSDONE-$LEG" >> $S/cg/done.log
