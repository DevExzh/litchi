#!/bin/bash
set -e
S=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0654
W=/home/zhuhe/code/litchi-worktrees/scratch-0654
CASES=xlsx_source_backed_cell_values_one_edit_save,xlsx_source_backed_cell_values_batch_edit_save,opc_source_overlay_one_part_save
cd /home/zhuhe/code/litchi
run() { # $1 = leg binary, $2 = label
  taskset -c 9 $W/bin/$1 --case $CASES --warmup 3 --samples 30 \
    --json $S/timing/$2.json > $S/timing/$2.txt 2>&1
  echo "leg $2 done"
}
run lpb-before A1
run lpb-after  B1
run lpb-after  B2
run lpb-before A2
run lpb-before A3
run lpb-before A4
