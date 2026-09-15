#!/usr/bin/env bash
set -u
S=$SCRATCH
A=/home/zhuhe/code/litchi-worktrees/targets/0606-before/release/litchi-perf-baseline
B=/home/zhuhe/code/litchi-worktrees/targets/0606-after/release/litchi-perf-baseline
CASES=ppt_semantic_open,ppt_semantic_list_slides,ppt_semantic_full_text,ppt_semantic_one_edit_save
FIX=/home/zhuhe/code/litchi/test-data/poi/test-data/slideshow/45543.ppt
# Harness, registered selectors, A1 B1 B2 A2 then two more A legs for the floor.
for pair in A1:$A B1:$B B2:$B A2:$A AA1:$A AA2:$A; do
  label=${pair%%:*}; bin=${pair#*:}
  taskset -c 26 "$bin" --case "$CASES" --samples 40 --warmup 5 \
    --json "$S/timing/$label.json" > /dev/null 2>"$S/timing/$label.err"
  echo "rc=$? $label" >> "$S/timing/done.log"
done
# Real fixture, whole operation, default and pinned glibc thresholds.
for cfg in default heap; do
  if [ $cfg = heap ]; then export MALLOC_MMAP_THRESHOLD_=8388608 MALLOC_TRIM_THRESHOLD_=8388608
  else unset MALLOC_MMAP_THRESHOLD_ MALLOC_TRIM_THRESHOLD_; fi
  for mode in eager-open eager-slides eager-text owned-edit-save; do
    for pair in A1:before B1:after B2:after A2:before AA1:before AA2:before; do
      label=${pair%%:*}; leg=${pair#*:}
      bin=/home/zhuhe/code/litchi-worktrees/targets/0606-drv-$leg/release/ppt0606
      taskset -c 26 "$bin" time "$mode" "$FIX" 5 40 \
        > "$S/timing/fx-$cfg-$mode-$label.txt" 2>/dev/null
      echo "rc=$? fx-$cfg-$mode-$label" >> "$S/timing/done.log"
    done
  done
done
echo TIMINGFINALDONE >> "$S/timing/done.log"
