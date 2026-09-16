#!/bin/bash
# Part (a) reproduction: does 0592 still show +5.49% on docx_file_eager_paragraph_count
# at base c7326f680?  Leg A = 0592 reverted (eager index), leg B = base (lazy index).
# Order A1 B1 B2 A2 on CPU 18.  A1 vs A2 is the A/A floor.
set -u
SP=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0643
ROOT=/home/zhuhe/code/litchi-worktrees/targets/0643-fsroot
FS="docx_file_eager_paragraph_count,docx_file_eager_full_text,docx_file_source_full_text"
mkdir -p $SP/out/repro
cd $ROOT || exit 1
for run in A1 B1 B2 A2; do
  case $run in
    A*) BIN=$SP/bin/pb-rev0592 ;;
    B*) BIN=$SP/bin/pb-base ;;
  esac
  taskset -c 18 $BIN --case $FS --filesystem-cache warm --filesystem-root $ROOT \
    --warmup 3 --samples 60 --json $SP/out/repro/fs-$run.json > /dev/null 2>$SP/out/repro/fs-$run.log
  echo "fs-$run exit=$?"
done
