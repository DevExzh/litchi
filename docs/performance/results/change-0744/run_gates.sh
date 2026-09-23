#!/usr/bin/env bash
# Change 0744 gates, run in the branch worktree with its own target directory.
set -u
cd /home/zhuhe/code/litchi-worktrees/0744-xlsx-eager-workbook-cell-path
export CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0744
export CARGO_BUILD_JOBS=6
LOG=/home/zhuhe/code/litchi-worktrees/scratch/0744/gates
: > "$LOG/gates.txt"
gate() { # name command...
  local name=$1; shift
  "$@" > "$LOG/$name.log" 2>&1
  local status=$?
  printf '%s\texit=%s\t%s\n' "$name" "$status" "$*" >> "$LOG/gates.txt"
}
echo "HEAD $(git rev-parse HEAD)" >> "$LOG/gates.txt"
gate fmt cargo fmt --all --check
gate check-xlsx cargo check -p litchi-xlsx --all-targets --locked --offline
gate check-xlsx-all-features cargo check -p litchi-xlsx --all-features --all-targets --locked --offline
gate check-facade cargo check -p litchi --features doc,docx,ppt,pptx,xls,xlsx,xlsb,odt --all-targets --locked --offline
gate clippy-lib cargo clippy -p litchi-xlsx --lib --no-deps --locked --offline -- -D warnings
gate clippy-all-targets cargo clippy -p litchi-xlsx --all-targets --no-deps --locked --offline -- -D warnings
gate test-xlsx cargo test -p litchi-xlsx --locked --offline
gate test-xlsx-all-features cargo test -p litchi-xlsx --all-features --locked --offline
gate test-facade cargo test -p litchi --features doc,docx,ppt,pptx,xls,xlsx,xlsb,odt --locked --offline
gate doc env RUSTDOCFLAGS="-D warnings" cargo doc -p litchi-xlsx --no-deps --locked --offline
gate crate-boundaries python3 tools/check_crate_boundaries.py
gate non-iwork python3 tools/non_iwork_gate.py verify
gate perf-claims python3 tools/check_perf_claims.py --registry docs/performance/claim-registry-v1.json --repo-root . --mode structural
echo done >> "$LOG/gates.txt"
