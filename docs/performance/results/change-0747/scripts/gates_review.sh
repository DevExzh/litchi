#!/bin/bash
# Change 0747 review-fix gates, fresh CARGO_TARGET_DIR under targets/0747.
set -u
cd /home/zhuhe/code/litchi-worktrees/0747-xlsx-publication-audit-reuse
export CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0747 CARGO_BUILD_JOBS=6
LOG=/home/zhuhe/code/litchi-worktrees/scratch/0747/gates2.txt
mkdir -p /home/zhuhe/code/litchi-worktrees/scratch/0747/gate2-logs
: > "$LOG"
gate() {
  local name=$1; shift
  local full=/home/zhuhe/code/litchi-worktrees/scratch/0747/gate2-logs/$name.log
  echo "=== $name" >> "$LOG"
  echo "\$ $*" >> "$LOG"
  "$@" > "$full" 2>&1
  local code=$?
  grep -E "^test result" "$full" | sort | uniq -c >> "$LOG"
  local total
  total=$(grep -E "^test result" "$full" | sed -E 's/.* ([0-9]+) passed; ([0-9]+) failed; ([0-9]+) ignored.*/\1 \2 \3/' | awk '{p+=$1; f+=$2; i+=$3} END {if (NR) print "total: " p " passed, " f " failed, " i " ignored, in " NR " test binaries"}')
  [ -n "$total" ] && echo "$total" >> "$LOG"
  grep -E "^error|FAILED|panicked|^warning" "$full" | head -20 >> "$LOG"
  tail -3 "$full" >> "$LOG"
  echo "exit=$code" >> "$LOG"
  echo >> "$LOG"
}
echo "HEAD $(git rev-parse HEAD) plus the uncommitted review fixes (diff hash $(git diff | sha256sum | cut -c1-16))" >> "$LOG"
gate fmt cargo fmt --all --check
gate clippy-lib cargo clippy -p xml-minifier -p litchi-opc -p litchi-xlsx --lib --no-deps --locked --offline -- -D warnings
gate clippy-all-targets cargo clippy -p xml-minifier -p litchi-opc -p litchi-xlsx --all-targets --no-deps --locked --offline -- -D warnings
gate test cargo test -p xml-minifier -p litchi-opc -p litchi-xlsx --locked --offline
RUSTDOCFLAGS="-D warnings" gate doc cargo doc -p xml-minifier -p litchi-opc -p litchi-xlsx --no-deps --locked --offline
echo "gates done" >> "$LOG"
