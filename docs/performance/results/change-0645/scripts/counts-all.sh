#!/bin/bash
set -u
S=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0645
for case in pptx_slide_move_boundary_save pptx_slide_remove_boundary_save pptx_eager_batch_edit_save pptx_eager_multi_slide_batch_edit_save; do
  for samples in 1 3; do
    for leg in before after; do
      echo "=== $leg $case s$samples ==="
      $S/cg-run.sh "$leg" "$S/stage/$leg" "$case" "$samples"
    done
  done
done
echo ALLDONE
