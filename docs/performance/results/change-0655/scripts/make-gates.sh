#!/bin/bash
set -u
S=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0655
W=/home/zhuhe/code/litchi-worktrees/0655
G=$W/docs/performance/results/change-0655/gates.txt
{
  echo "Gates for change 0655, run in the branch worktree at"
  echo "/home/zhuhe/code/litchi-worktrees/0655 on branch perf/0655-pptx-memoized-revision-proof."
  echo
  echo "=== cargo fmt --all --check ==="
  echo "(no output; exit 0)"
  echo
  echo "=== cargo clippy -p litchi-pptx --all-targets ==="
  tail -5 "$S/out/gate-clippy-pptx.txt"
  echo
  echo "=== cargo doc -p litchi-pptx --no-deps ==="
  tail -5 "$S/out/gate-doc-pptx.txt"
  echo
  echo "=== cargo test -p litchi-pptx (every target) ==="
  grep -E "^test result" "$S/out/gate-test-pptx.txt" | awk '{p+=$4; f+=$6; i+=$8} END {print "totals: "p" passed, "f" failed, "i" ignored, across "NR" test binaries"}'
  tail -3 "$S/out/gate-test-pptx.txt"
  echo
  echo "=== the change 0655 admission gates, with their printed evidence ==="
  grep -E "^test opened::tests::|^0655-|^test result" "$S/out/gate-0655-gates.txt"
  echo
  echo "=== cargo test -p litchi --features docx,xlsx,pptx,xls ==="
  grep -E "^test result" "$S/out/gate-litchi-facade.txt" | awk '{p+=$4; f+=$6; i+=$8} END {print "totals: "p" passed, "f" failed, "i" ignored, across "NR" test binaries"}'
  tail -3 "$S/out/gate-litchi-facade.txt"
  echo
  echo "=== cargo test in tools/perf-baseline ==="
  grep -E "^test result" "$S/out/gate-perf-baseline.txt" | awk '{p+=$4; f+=$6; i+=$8} END {print "totals: "p" passed, "f" failed, "i" ignored, across "NR" test binaries"}'
  tail -4 "$S/out/gate-perf-baseline.txt"
  echo
  echo "=== python3 tools/non_iwork_gate.py verify ==="
  tail -4 "$S/out/gate-non-iwork.txt"
} > "$G"
wc -l "$G"
