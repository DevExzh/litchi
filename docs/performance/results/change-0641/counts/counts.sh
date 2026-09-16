#!/bin/bash
# The control for change 0641: logical reads, read bytes, `version()` calls and
# the harness's own observation, for every scenario on both legs.
#
# Five samples per run, one warmup. The harness reports per-sample counters; the
# fold asserts the five samples of a run agree with each other before comparing
# the legs, so a scenario whose counters were not deterministic would fail here
# rather than be averaged.
set -e
S=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0641
cd /home/zhuhe/code/litchi-worktrees/0641
while IFS='|' read -r label file idx op; do
  [ -z "$label" ] && continue
  for leg in before after; do
    "$S/bin/xsa-$leg" --input "$file" --operation "$op" --worksheet-index "$idx" \
      --warmups 1 --samples 5 > "$S/counts/$label-$leg.json"
  done
done < "$1"
