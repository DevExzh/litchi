#!/usr/bin/env bash
# Confirmation window for change 0637's two eager PPTX filesystem selectors,
# with two controls this change cannot reach: `docx_file_eager_full_text`
# (litchi-docx) and `opc_file_eager_open` (litchi-opc). If the controls move by
# the same amount as the PPTX selectors, the shift is a whole-binary effect,
# not the memo.
set -euo pipefail
S="$1"; CPU=13
A="$S/bin/litchi-perf-baseline-before"
B="$S/bin/litchi-perf-baseline-after"
CWD=/home/zhuhe/code/litchi-worktrees/before-c7326f680/tools/perf-baseline
CASES=pptx_file_eager_slide_count,pptx_file_eager_selected_slide,docx_file_eager_full_text,opc_file_eager_open
mkdir -p "$S/timing/raw-repeat"
date -Is >> "$S/timing/window.txt"; uptime >> "$S/timing/window.txt"
for run in A1 B1 B2 A2; do
  case $run in A1|A2) bin="$A" ;; *) bin="$B" ;; esac
  ( cd "$CWD" && taskset -c $CPU "$bin" --case "$CASES" --samples 100 --warmup 5 \
      --filesystem-cache warm > "$S/timing/raw-repeat/$run.json" 2>/dev/null )
  uptime >> "$S/timing/window.txt"
done
date -Is >> "$S/timing/window.txt"; uptime >> "$S/timing/window.txt"
