#!/usr/bin/env bash
# Gate runner for change 0593. Each section prints its own tail.
set -u
cd /home/zhuhe/code/litchi-worktrees/0593 || exit 1
SC=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0593
mkdir -p "$SC/gates"

section() {
  echo
  echo "### $*"
}

{
  section "cargo fmt --all --check"
  nice -n 5 cargo fmt --all --check 2>&1 | tail -5
  echo "exit=${PIPESTATUS[0]}"

  section "cargo clippy -p litchi-opc --all-targets"
  nice -n 5 cargo clippy -p litchi-opc --all-targets 2>&1 | tail -6
  echo "exit=${PIPESTATUS[0]}"

  section "cargo test -p litchi-opc"
  nice -n 5 cargo test -p litchi-opc -j 6 2>&1 | grep -E "^(test result|error|running|warning)" | tail -50
  echo "exit=${PIPESTATUS[0]}"

  section "cargo doc -p litchi-opc --no-deps"
  nice -n 5 cargo doc -p litchi-opc --no-deps 2>&1 | tail -6
  echo "exit=${PIPESTATUS[0]}"

  section "cargo clippy -p litchi-xlsx -p litchi-docx -p litchi-pptx --all-targets"
  nice -n 5 cargo clippy -p litchi-xlsx -p litchi-docx -p litchi-pptx --all-targets -j 6 2>&1 | tail -6
  echo "exit=${PIPESTATUS[0]}"

  section "cargo test -p litchi-xlsx -p litchi-docx -p litchi-pptx"
  nice -n 5 cargo test -p litchi-xlsx -p litchi-docx -p litchi-pptx -j 6 2>&1 | grep -E "^(test result|error)" | tail -60
  echo "exit=${PIPESTATUS[0]}"
} > "$SC/gates/gates.txt" 2>&1
echo "gates finished"
