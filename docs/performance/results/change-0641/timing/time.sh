#!/bin/bash
# Paired timing for change 0641, ordered A1 B1 B2 A2 on one pinned CPU, with an
# A/A leg (before against before) in the same window as the floor.
set -e
S=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0641
CPU=24
N=${N:-40}
WARM=${WARM:-8}
cd /home/zhuhe/code/litchi-worktrees/0641
run() { # binary file idx op outfile
  taskset -c $CPU "$1" --input "$2" --worksheet-index "$3" --operation "$4" \
    --warmups $WARM --samples $N 2>/dev/null \
    | python3 -c 'import json,sys; [print(v) for v in json.load(sys.stdin)["elapsed_samples_ns"]]' > "$5"
}
while IFS='|' read -r label file idx op; do
  [ -z "$label" ] && continue
  run "$S/bin/xsa-before" "$file" "$idx" "$op" "$S/timing/$label-A1.txt"
  run "$S/bin/xsa-after"  "$file" "$idx" "$op" "$S/timing/$label-B1.txt"
  run "$S/bin/xsa-after"  "$file" "$idx" "$op" "$S/timing/$label-B2.txt"
  run "$S/bin/xsa-before" "$file" "$idx" "$op" "$S/timing/$label-A2.txt"
  run "$S/bin/xsa-before" "$file" "$idx" "$op" "$S/timing/$label-A3.txt"
  echo "timing done $label"
done < "$1"
