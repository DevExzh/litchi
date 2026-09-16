#!/bin/bash
# change 0656: callgrind isolation pairs on CPU 11, serial.
set -u
SCR=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0656
for CASE in pptx_cross_copy_plain_lifecycle pptx_cross_copy_media_rich_lifecycle; do
  for LEG in before after; do
    for S in 1 3; do
      echo "=== $LEG $CASE s$S : $(date -Is) ==="
      "$SCR/scripts/cg.sh" "$SCR/bin/litchi-perf-baseline-$LEG" "$LEG" "$CASE" "$S" "$SCR/counts"
    done
  done
done
echo "ALL COUNTS DONE $(date -Is)"
