#!/bin/bash
# Paired timing driver for change 0636. Legs run A1 B1 B2 A2 on one pinned CPU.
set -e
S=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0636
CPU=12
N=${N:-40}
WARM=${WARM:-5}
cd /home/zhuhe/code/litchi-worktrees/0636
run() { # binary kind op file tag
  taskset -c $CPU "$1" time "$2" "$3" $((N+WARM)) "$4" | tail -n +$((WARM+1)) | cut -f1 > "$S/$5"
}
for spec in "$@"; do
  IFS='|' read -r kind op file label <<< "$spec"
  run "$S/bin/probe-before" "$kind" "$op" "$file" "t-$label-A1.txt"
  run "$S/bin/probe-after"  "$kind" "$op" "$file" "t-$label-B1.txt"
  run "$S/bin/probe-after"  "$kind" "$op" "$file" "t-$label-B2.txt"
  run "$S/bin/probe-before" "$kind" "$op" "$file" "t-$label-A2.txt"
  echo "done $label"
done
