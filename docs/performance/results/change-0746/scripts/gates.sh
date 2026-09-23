#!/usr/bin/env bash
# Change 0746 gates, run in the worktree with the change's own target dir.
cd /home/zhuhe/code/litchi-worktrees/0746-xls-validation-only-parse || exit 1
export CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0746
export CARGO_BUILD_JOBS=6
OUT=/home/zhuhe/code/litchi-worktrees/0746-xls-validation-only-parse/docs/performance/results/change-0746/gates.txt
: > "$OUT"
run() {
  echo "\$ $*" >> "$OUT"
  "$@" > /home/zhuhe/code/litchi-worktrees/scratch/0746/gate-step.log 2>&1
  status=$?
  tail -n 4 /home/zhuhe/code/litchi-worktrees/scratch/0746/gate-step.log | sed 's/^/    /' >> "$OUT"
  echo "exit $status" >> "$OUT"
  echo >> "$OUT"
}
echo "HEAD $(git rev-parse HEAD)" >> "$OUT"
echo "rustc $(rustc --version)" >> "$OUT"
echo >> "$OUT"
run cargo fmt --all --check
run cargo check -p litchi-xls --all-targets --locked --offline
run cargo check -p litchi --features doc,docx,ppt,pptx,xls,xlsx,xlsb,odt --all-targets --locked --offline
run cargo clippy -p litchi-xls --lib --no-deps --locked --offline -- -D warnings
run cargo clippy -p litchi-xls --all-targets --no-deps --locked --offline -- -D warnings
run cargo clippy -p litchi-xls --all-targets --no-deps --locked --offline -- -D warnings -A clippy::unusual_byte_groupings
run cargo test -p litchi-xls --locked --offline
run cargo test --release -p litchi-xls --lib --locked --offline -- generic_commit_publishes validation_only_tests deterministically
run cargo test -p litchi --features doc,docx,ppt,pptx,xls,xlsx,xlsb,odt --locked --offline
env RUSTDOCFLAGS="-D warnings" cargo doc -p litchi-xls --no-deps --locked --offline > /home/zhuhe/code/litchi-worktrees/scratch/0746/gate-step.log 2>&1; s=$?; echo '$ RUSTDOCFLAGS="-D warnings" cargo doc -p litchi-xls --no-deps --locked --offline' >> "$OUT"; tail -n 3 /home/zhuhe/code/litchi-worktrees/scratch/0746/gate-step.log | sed 's/^/    /' >> "$OUT"; echo "exit $s" >> "$OUT"; echo >> "$OUT"
run python3 tools/check_crate_boundaries.py
run python3 tools/non_iwork_gate.py verify
run python3 tools/check_perf_claims.py --registry docs/performance/claim-registry-v1.json --repo-root . --mode structural
echo "note: the --all-targets clippy exit 101 is the pre-existing unusual_byte_groupings error in crates/litchi-xls/tests/xls_query_index_cache.rs:546 (file untouched by this change); the identical command on the base checkout 009d515bef also exits 101 with the same single error." >> "$OUT"
echo "gates done" >> "$OUT"
