#!/usr/bin/env bash
# Gate runner for change 0618. Each section prints its own tail.
set -u
cd /home/zhuhe/code/litchi-worktrees/0618 || exit 1
SC=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0618
mkdir -p "$SC/gates"
section() { echo; echo "### $*"; }
{
  section "cargo fmt --all --check"
  nice -n 5 cargo fmt --all --check 2>&1 | tail -5
  echo "exit=${PIPESTATUS[0]}"

  section "cargo clippy -p soapberry-zip --all-targets"
  nice -n 5 cargo clippy -p soapberry-zip --all-targets -j 8 2>&1 | tail -6
  echo "exit=${PIPESTATUS[0]}"

  section "cargo test -p soapberry-zip"
  nice -n 5 cargo test -p soapberry-zip -j 8 2>&1 | grep -E "^(test result|error|warning: unused)" | tail -30
  echo "exit=${PIPESTATUS[0]}"

  section "cargo doc -p soapberry-zip --no-deps"
  nice -n 5 cargo doc -p soapberry-zip --no-deps -j 8 2>&1 | tail -6
  echo "exit=${PIPESTATUS[0]}"

  section "cargo test -p litchi-opc (consumer: the OOXML publication path)"
  nice -n 5 cargo test -p litchi-opc -j 8 2>&1 | grep -E "^(test result|error)" | tail -30
  echo "exit=${PIPESTATUS[0]}"

  section "cargo test -p litchi-xlsx -p litchi-docx -p litchi-pptx (OOXML editors)"
  nice -n 5 cargo test -p litchi-xlsx -p litchi-docx -p litchi-pptx -j 8 2>&1 | grep -E "^(test result|error)" | tail -50
  echo "exit=${PIPESTATUS[0]}"

  section "cargo test -p litchi-odt -p litchi-odc -p litchi-odf-common -p litchi-iwa-archive (other streaming-writer consumers)"
  nice -n 5 cargo test -p litchi-odt -p litchi-odc -p litchi-odf-common -p litchi-iwa-archive -j 8 2>&1 | grep -E "^(test result|error)" | tail -50
  echo "exit=${PIPESTATUS[0]}"
} > "$SC/gates/gates.txt" 2>&1
echo "gates finished"
