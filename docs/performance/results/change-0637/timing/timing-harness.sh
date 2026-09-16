#!/usr/bin/env bash
# Paired timing of the three harness selectors named in change 0637's brief.
# A = before leg (c7326f680), B = after leg. Legs run A1 B1 B2 A2 so the A/A
# pair brackets the B pair and both see the same window of host load.
# Both binaries are staged outside any Cargo target directory (0627).
set -euo pipefail
S="$1"; CPU=13
A="$S/bin/litchi-perf-baseline-before"
B="$S/bin/litchi-perf-baseline-after"
CWD=/home/zhuhe/code/litchi-worktrees/before-c7326f680/tools/perf-baseline
SAMPLES=${SAMPLES:-40}
leg() { # leg <name> <binary> <case-list> <extra args...>
  local name="$1" bin="$2" cases="$3"; shift 3
  ( cd "$CWD" && taskset -c $CPU "$bin" --case "$cases" --samples "$SAMPLES" --warmup 5 "$@" \
      > "$S/timing/raw/$name.json" 2>/dev/null )
}
mkdir -p "$S/timing/raw"
CASES=pptx_file_eager_slide_count,pptx_file_eager_selected_slide,pptx_semantic_full_text
date -Is > "$S/timing/window.txt"; uptime >> "$S/timing/window.txt"
for run in A1 B1 B2 A2; do
  case $run in A1|A2) bin="$A" ;; *) bin="$B" ;; esac
  leg "$run" "$bin" "$CASES" --filesystem-cache warm
  uptime >> "$S/timing/window.txt"
done
date -Is >> "$S/timing/window.txt"; uptime >> "$S/timing/window.txt"
