#!/bin/bash
# Change 0750: exact instructions of one audit, from the audit probe at a
# fixed iteration count against zero iterations (the difference cancels
# loading the parts), for both legs.
set -u
S=/home/zhuhe/code/litchi-worktrees/scratch/0750
mkdir -p $S/cg-probe && cd $S/cg-probe
for leg in before after; do
  if [ $leg = before ]; then B=/home/zhuhe/code/litchi-worktrees/targets/0750-before/release/litchi-audit-probe-0750
  else B=/home/zhuhe/code/litchi-worktrees/targets/0750-census-after/release/litchi-audit-probe-0750; fi
  for spec in "source/ws-structured|4" "source/ws-patriarch|2" "source/docx-drawing|20" "source/docx-table-alignment|100" "source/pptx-slide11|40" "pair/ws-structured|4" "source/corpus-accepted|1" "authored/ws-patriarch|2" "authored/tiny|10000" "source/tiny|10000"; do
    case_=${spec%%|*}; n=${spec#*|}; tag=$(echo $case_ | tr '/' '_')
    for it in 0 $n; do
      taskset -c 24 valgrind --tool=callgrind --callgrind-out-file=cg-$leg-$tag-$it.out $B $S/parts $case_ --iterations $it > /dev/null 2>&1
    done
    echo "$leg $case_ $n" >> status.txt
  done
done
echo done >> status.txt
