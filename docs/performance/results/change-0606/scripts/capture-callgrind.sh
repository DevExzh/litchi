#!/usr/bin/env bash
# Callgrind isolation pairs (s=10 vs s=110; per-op = (Ir110-Ir10)/100), pinned to CPU 26.
set -u
S=$SCRATCH
R=/home/zhuhe/code/litchi
export TMPDIR=$S/tmp RAYON_NUM_THREADS=1
mkdir -p "$TMPDIR" "$S/cg"
FIX=$R/test-data/poi/test-data/slideshow/45543.ppt
for leg in before after; do
  B=/home/zhuhe/code/litchi-worktrees/targets/0606-drv-$leg/release/ppt0606
  for mode in eager-open eager-slides eager-text owned-edit-save; do
    for n in 10 110; do
      out=$S/cg/$leg-$mode-s$n
      setarch x86_64 -R taskset -c 26 valgrind --tool=callgrind \
        --callgrind-out-file=$out.out --cache-sim=no --branch-sim=no \
        "$B" profile "$mode" "$FIX" 1 "$n" > $out.stdout 2> $out.stderr
      echo "rc=$? $leg-$mode-s$n" >> $S/cg/done.log
    done
  done
done
echo ALLDONE >> $S/cg/done.log
