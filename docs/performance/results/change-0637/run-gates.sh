#!/usr/bin/env bash
# Gates for change 0637, run in the change's own worktree.
#
# litchi-pptx is the only crate edited. `litchi` is its consumer through the
# presentation facade, and `tools/perf-baseline` is a separate Cargo project
# that links litchi-pptx and is reachable from no per-crate gate.
set -u
W="${1:-/home/zhuhe/code/litchi-worktrees/0637}"
section() { printf '\n===== %s =====\n' "$1"; }

section "cargo fmt --all --check"
( cd "$W" && cargo fmt --all --check ) 2>&1 | tail -5; echo "exit=${PIPESTATUS[0]}"

section "cargo clippy -p litchi-pptx --all-targets"
( cd "$W" && cargo clippy -p litchi-pptx --all-targets ) 2>&1 | tail -5; echo "exit=${PIPESTATUS[0]}"

section "cargo test -p litchi-pptx"
( cd "$W" && cargo test -p litchi-pptx ) 2>&1 | grep -E "^test result|^error|warning: unused" | tail -40; echo "exit=${PIPESTATUS[0]}"

section "cargo doc -p litchi-pptx --no-deps"
( cd "$W" && cargo doc -p litchi-pptx --no-deps ) 2>&1 | tail -5; echo "exit=${PIPESTATUS[0]}"

section "cargo test -p litchi --features docx,xlsx,pptx,xls"
( cd "$W" && cargo test -p litchi --features docx,xlsx,pptx,xls ) 2>&1 | grep -E "^test result|^error" | tail -40; echo "exit=${PIPESTATUS[0]}"

section "cargo test (tools/perf-baseline)"
( cd "$W/tools/perf-baseline" && cargo test ) 2>&1 | grep -E "^test result|^error" | tail -20; echo "exit=${PIPESTATUS[0]}"
