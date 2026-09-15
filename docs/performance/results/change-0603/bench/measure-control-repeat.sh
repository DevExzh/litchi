#!/bin/bash
# Second interleaved pass over the marker-free control (A3 B3 B4 A4), so the
# control's own before/after difference rests on four legs of each rather than
# two. Same binaries, same CPU, run immediately after the main window.
set -u
S=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0603
A=/home/zhuhe/code/litchi-worktrees/targets/0603-before/release/xlsx-admission-probe
B=/home/zhuhe/code/litchi-worktrees/targets/0603-after/release/xlsx-admission-probe
for spec in A3:$A B3:$B B4:$B A4:$A; do
  name=${spec%%:*}; bin=${spec#*:}
  taskset -c 24 "$bin" bench "$S/derived/no_drawing_patriarch-proj-plain.xlsx" "Лист 1" A1 5 30 \
    > "$S/bench/legs/ndp-$name.txt" 2>&1
done
echo REPEATDONE
