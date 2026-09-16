#!/usr/bin/env bash
# Paired timing for change 0634.  A = before, B = after; AA1/AA2 are two more
# before legs in the same window for the A/A floor.  usage: capture-timing.sh <cpu>
set -u
CPU=$1
S=${SCRATCH:?}
R=/home/zhuhe/code/litchi
FIX=$R/test-data/poi/test-data/slideshow/45543.ppt
NFIX=$R/test-data/poi/test-data/slideshow/headers_footers_2007.ppt
A=$S/bin/litchi-perf-baseline-before
B=$S/bin/litchi-perf-baseline-after
CASES=ppt_semantic_open,ppt_semantic_list_slides,ppt_semantic_full_text,ppt_semantic_one_edit_save
export TMPDIR=$S/tmp RAYON_NUM_THREADS=1
mkdir -p "$TMPDIR" "$S/timing"
: > "$S/timing/done.log"
for pair in A1:$A B1:$B B2:$B A2:$A AA1:$A AA2:$A; do
  label=${pair%%:*}; bin=${pair#*:}
  taskset -c $CPU "$bin" --case "$CASES" --samples 40 --warmup 5 \
    --json "$S/timing/$label.json" > /dev/null 2>"$S/timing/$label.err"
  echo "rc=$? $label" >> "$S/timing/done.log"
done
for cfg in default heap; do
  if [ $cfg = heap ]; then export MALLOC_MMAP_THRESHOLD_=8388608 MALLOC_TRIM_THRESHOLD_=8388608
  else unset MALLOC_MMAP_THRESHOLD_ MALLOC_TRIM_THRESHOLD_; fi
  for spec in eager-open:$FIX eager-slides:$FIX eager-text:$FIX eager-notes:$NFIX owned-edit-save:$FIX; do
    mode=${spec%%:*}; fx=${spec#*:}
    for pair in A1:before B1:after B2:after A2:before AA1:before AA2:before; do
      label=${pair%%:*}; leg=${pair#*:}
      taskset -c $CPU "$S/bin/ppt0634-$leg" time "$mode" "$fx" 5 40 \
        > "$S/timing/fx-$cfg-$mode-$label.txt" 2>/dev/null
      echo "rc=$? fx-$cfg-$mode-$label" >> "$S/timing/done.log"
    done
  done
done
echo TIMINGDONE >> "$S/timing/done.log"
