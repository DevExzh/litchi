#!/usr/bin/env bash
# git bisect run script for change 0629.
# Verdict for one commit: does the litchi facade test
# managed_docx_facade_paragraph_text_avoids_rich_paragraph_refusal pass?
# exit 0 = good (test ran and passed), 1 = bad (test ran and failed),
# 125 = skip (does not build, or the test is not present at this commit).
set -u
WT=/home/zhuhe/code/litchi-worktrees/0629-bisect
TD=/home/zhuhe/code/litchi-worktrees/targets/0629-bisect
LOG=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0629/bisect-log.txt
TEST=managed_docx_facade_paragraph_text_avoids_rich_paragraph_refusal
cd "$WT" || exit 125
cp -f /home/zhuhe/code/litchi/Cargo.lock "$WT/Cargo.lock" 2>/dev/null
SHA=$(git rev-parse --short HEAD)
OUT=$(CARGO_TARGET_DIR="$TD" taskset -c 23 cargo test -p litchi --offline --features docx "$TEST" 2>&1)
RC=$?
if ! printf '%s' "$OUT" | grep -q "test document::doc::tests::$TEST \.\.\."; then
  # test not compiled/not present, or the build failed: skip
  echo "$SHA SKIP rc=$RC $(git log -1 --format=%s HEAD)" >> "$LOG"
  printf '%s\n' "$OUT" | tail -25 >> "$LOG"
  exit 125
fi
if printf '%s' "$OUT" | grep -q "test document::doc::tests::$TEST ... ok"; then
  echo "$SHA GOOD $(git log -1 --format=%s HEAD)" >> "$LOG"
  exit 0
fi
echo "$SHA BAD $(git log -1 --format=%s HEAD)" >> "$LOG"
printf '%s' "$OUT" | grep -A3 "panicked at" | head -8 >> "$LOG"
exit 1
