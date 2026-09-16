#!/bin/bash
set -u
S=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0645
for case in pptx_eager_batch_edit_save pptx_eager_multi_slide_batch_edit_save pptx_slide_move_boundary_save pptx_slide_remove_boundary_save; do
  echo "=== timing $case ==="
  $S/time-run.sh "$case" 30 3
  python3 $S/time-report.py "$case" | tee $S/out/time-$case.txt
  python3 $S/phase-report.py "$case" > $S/out/phases-$case.txt 2>&1 || true
done
for case in pptx_eager_batch_edit_save pptx_eager_multi_slide_batch_edit_save pptx_slide_move_boundary_save pptx_slide_remove_boundary_save; do
  echo "=== perf $case ==="
  $S/perf-run.sh "$case" 30 3
  python3 $S/perf-report.py "$case" | tee $S/out/perf-$case.txt
done
echo TIMING-DONE
