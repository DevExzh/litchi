#!/usr/bin/env bash
# Run one gate command in the 0753 worktree and append it, its exit code and
# the tail of its output to gates.txt. usage: gate.sh "command..."
W=/home/zhuhe/code/litchi-worktrees/0753-legacy-fresh-writer-text-paths
G=/home/zhuhe/code/litchi-worktrees/scratch/0753/results/gates.txt
export CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0753 CARGO_BUILD_JOBS=6 TMPDIR=/home/zhuhe/code/litchi-worktrees/scratch/0753/tmp
cd "$W"
LOG=$(mktemp -p /home/zhuhe/code/litchi-worktrees/scratch/0753/tmp gate.XXXXXX)
bash -c "$1" > "$LOG" 2>&1
code=$?
{
  echo "\$ $1"
  if grep -q "^test result" "$LOG"; then
    grep -E "^test result" "$LOG" | awk '{p+=$4; f+=$6; i+=$8; n+=1} END {printf "    %d suites, %d passed, %d failed, %d ignored\n", n, p, f, i}'
    grep -E "FAILED|panicked" "$LOG" | head -5 | sed 's/^/    /'
  else
    tail -4 "$LOG" | sed 's/^/    /'
  fi
  echo "exit $code"
  echo
} >> "$G"
rm -f "$LOG"
echo "exit=$code"
