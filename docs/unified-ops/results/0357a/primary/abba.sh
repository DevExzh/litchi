#!/bin/bash
B=/home/zhuhe/code/litchi-worktrees/.coord/w3-ppt; cd $B/perf2; CPU=29
run() { taskset -c $CPU $B/tgt-$1/release/examples/$2 ${@:4} > $3 2>>log.txt || echo "FAIL $3" >> log.txt; }
echo "start $(date +%T) load $(cut -d' ' -f1 /proc/loadavg)" >> log.txt
for leg in before after after before; do
  n=$((n+1))
  run $leg ppt_paired_probe single-$leg-$n.json single 61
  run $leg ppt_paired_probe batch-$leg-$n.json batch 15
  run $leg ppt_elements_probe elements-$leg-$n.json 21
  echo "leg $n $leg $(date +%T) load $(cut -d' ' -f1 /proc/loadavg)" >> log.txt
done
echo done >> log.txt
