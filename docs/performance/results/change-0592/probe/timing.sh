#!/bin/bash
# Paired timing, order A1 B1 B2 A2 on CPU 12. A1 vs A2 is the A/A floor.
set -u
SP=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0592
ROOT=/home/zhuhe/code/litchi-worktrees/targets/0592-fsroot
SEM="docx_semantic_full_text,docx_semantic_one_paragraph,docx_semantic_list_paragraphs,docx_semantic_open"
FS="docx_file_eager_full_text,docx_file_eager_paragraph_count,docx_file_source_full_text,docx_file_eager_open_full_text_lifecycle"
mkdir -p $SP/out/timing
cd $ROOT || exit 1
for run in A1 B1 B2 A2; do
  case $run in
    A*) BIN=/home/zhuhe/code/litchi-worktrees/targets/0592-before/release/litchi-perf-baseline ;;
    B*) BIN=/home/zhuhe/code/litchi-worktrees/targets/0592-after/release/litchi-perf-baseline ;;
  esac
  taskset -c 12 $BIN --case $SEM --semantic-shape large --warmup 3 --samples 30 \
    --json $SP/out/timing/sem-$run.json > /dev/null 2>$SP/out/timing/sem-$run.log
  echo "sem-$run exit=$?"
  taskset -c 12 $BIN --case $FS --filesystem-cache warm --filesystem-root $ROOT \
    --warmup 3 --samples 30 --json $SP/out/timing/fs-$run.json > /dev/null 2>$SP/out/timing/fs-$run.log
  echo "fs-$run exit=$?"
done
