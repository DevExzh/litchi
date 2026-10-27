#!/bin/bash
# cg-list.sh <routes.tsv> <out.tsv> <sides> [jobs]
F=/home/zhuhe/code/litchi-worktrees/.coord/fin-perf
list=$1 out=$2 sides=$3 jobs=${4:-8}
while IFS=$'\t' read -r group route; do
  for side in $sides; do for leg in leaf unified; do echo "$side $group $route $leg"; done; done
done < "$list" | xargs -P "$jobs" -L 1 $F/cg-one2.sh >> "$out"
echo DONE >> "$out"
