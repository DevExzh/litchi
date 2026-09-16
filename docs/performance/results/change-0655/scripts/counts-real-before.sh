#!/bin/bash
set -u
S=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0655
for case in pptx_real_file_ordinary_save_edit pptx_real_file_ordinary_save_lifecycle; do
  for samples in 1 3; do
    echo "=== before $case s$samples ==="
    $S/cg-run-real.sh before "$S/stage/before" "$case" "$samples"
  done
done
echo REALBEFOREDONE
