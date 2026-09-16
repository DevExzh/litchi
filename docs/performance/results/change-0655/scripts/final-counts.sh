#!/bin/bash
set -u
S=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0655
until grep -q ALLDONE $S/out/counts-all.txt 2>/dev/null; do sleep 15; done
until [ -f /home/zhuhe/code/litchi-worktrees/targets/0655-after/release/litchi-perf-baseline ] \
   && [ /home/zhuhe/code/litchi-worktrees/targets/0655-after/release/litchi-perf-baseline -nt $S/stage/after ]; do sleep 15; done
cp /home/zhuhe/code/litchi-worktrees/targets/0655-after/release/litchi-perf-baseline $S/stage/after
echo "restaged after: $(sha256sum $S/stage/after)"
$S/counts-real-before.sh
$S/counts-after.sh
echo FINAL-COUNTS-DONE
