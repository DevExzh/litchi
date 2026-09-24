#!/usr/bin/env bash
# Gate runner for change 0760: every command runs in the worktree with its own
# target directory, --locked --offline, and TMPDIR under the change's scratch
# directory. Each command, its exit code and the tail of its output are
# appended to gates.txt.
set -u
WORKTREE=/home/zhuhe/code/litchi-worktrees/0760-pptx-slide-root-memo-reapply
OUT=${1:-/home/zhuhe/code/litchi-worktrees/scratch/0760/packet/gates.txt}
export CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0760
export CARGO_BUILD_JOBS=6
export TMPDIR=/home/zhuhe/code/litchi-worktrees/scratch/0760/tmp
cd "$WORKTREE" || exit 1
echo "# gates for change 0760 at $(git rev-parse HEAD) ($(date -u +%Y-%m-%dT%H:%MZ))" > "$OUT"
run() {
  local log
  log=$(mktemp "$TMPDIR/gate.XXXXXX")
  "$@" > "$log" 2>&1
  local code=$?
  {
    echo
    echo "## $*"
    echo "exit: $code"
    grep -E "^test result:" "$log" | awk '{p+=$4; f+=$6; i+=$8} END {if (NR) print "test totals: " p " passed, " f " failed, " i " ignored (" NR " test binaries)"}'
    echo '```'
    tail -n 12 "$log"
    echo '```'
  } >> "$OUT"
  rm -f "$log"
}
run cargo fmt --all --check
run cargo check -p litchi-pptx --all-targets --locked --offline
run cargo check -p litchi-pptx --all-targets --all-features --locked --offline
run cargo check -p litchi --features doc,docx,ppt,pptx,xls,xlsx,xlsb,odt --all-targets --locked --offline
run cargo clippy -p litchi-pptx --lib --no-deps --locked --offline -- -D warnings
run cargo clippy -p litchi-pptx --all-targets --no-deps --locked --offline -- -D warnings
run cargo test -p litchi-pptx --locked --offline
run cargo test -p litchi-pptx --all-features --locked --offline
run cargo test -p litchi --features doc,docx,ppt,pptx,xls,xlsx,xlsb,odt --locked --offline
run env RUSTDOCFLAGS="-D warnings" cargo doc -p litchi-pptx --no-deps --locked --offline
run python3 tools/check_crate_boundaries.py
run python3 tools/non_iwork_gate.py verify
run python3 tools/check_perf_claims.py --registry docs/performance/claim-registry-v1.json --repo-root . --mode structural
echo "done" >> "$OUT"
