#!/bin/bash
# Paired timing for the scenarios whose A/A floor showed drift inside one window.
# Three rounds of A1 B1 B2 A2 A3, pooled per leg, so a monotone drift in the
# window averages out instead of being charged to one leg.
set -e
S=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0641
CPU=24
N=${N:-20}
WARM=${WARM:-8}
ROUNDS=${ROUNDS:-3}
cd /home/zhuhe/code/litchi-worktrees/0641
run() { taskset -c $CPU "$1" --input "$2" --worksheet-index "$3" --operation "$4" \
    --warmups $WARM --samples $N 2>/dev/null \
    | python3 -c 'import json,sys; [print(v) for v in json.load(sys.stdin)["elapsed_samples_ns"]]' >> "$5"; }
while IFS='|' read -r label file idx op; do
  [ -z "$label" ] && continue
  for leg in A1 B1 B2 A2 A3; do : > "$S/timing/$label-$leg.txt"; done
  for round in $(seq 1 $ROUNDS); do
    run "$S/bin/xsa-before" "$file" "$idx" "$op" "$S/timing/$label-A1.txt"
    run "$S/bin/xsa-after"  "$file" "$idx" "$op" "$S/timing/$label-B1.txt"
    run "$S/bin/xsa-after"  "$file" "$idx" "$op" "$S/timing/$label-B2.txt"
    run "$S/bin/xsa-before" "$file" "$idx" "$op" "$S/timing/$label-A2.txt"
    run "$S/bin/xsa-before" "$file" "$idx" "$op" "$S/timing/$label-A3.txt"
  done
  echo "rounds done $label"
done < "$1"
