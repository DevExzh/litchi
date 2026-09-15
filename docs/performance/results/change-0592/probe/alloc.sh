#!/bin/bash
# Allocation isolation pairs: N=4 and N=20, per-operation = (T20 - T4) / 16.
set -u
SP=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0592
OPS="eager_document eager_text eager_write_text eager_paragraph_first eager_paragraph_repeat8 eager_paragraph_count eager_paragraphs eager_tables source_document source_text source_paragraph_first source_paragraph_repeat8 source_paragraph_count"
printf "%-26s %-6s %-7s %14s %16s\n" operation shape leg allocations allocated_bytes
for shape in 200 10000; do
  for op in $OPS; do
    for leg in before after; do
      BIN=/home/zhuhe/code/litchi-worktrees/targets/0592-probe-$leg/release/probe0592
      lo=$(taskset -c 12 $BIN $op $shape 4  | sed 's/.*allocations=\([0-9]*\) allocated_bytes=\([0-9]*\)/\1 \2/')
      hi=$(taskset -c 12 $BIN $op $shape 20 | sed 's/.*allocations=\([0-9]*\) allocated_bytes=\([0-9]*\)/\1 \2/')
      a=$(( ( $(echo $hi | cut -d' ' -f1) - $(echo $lo | cut -d' ' -f1) ) / 16 ))
      b=$(( ( $(echo $hi | cut -d' ' -f2) - $(echo $lo | cut -d' ' -f2) ) / 16 ))
      printf "%-26s %-6s %-7s %14d %16d\n" $op $shape $leg $a $b
    done
  done
done
