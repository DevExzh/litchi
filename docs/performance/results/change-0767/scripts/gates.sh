#!/bin/bash
# Change 0767 gates, run from the worktree root in the foreground order below.
# Each command's exit code is appended to gates.log; a failure does not stop
# the later gates.
set -u
cd /home/zhuhe/code/litchi-worktrees/0767-cfb-reparse-linear
export RUSTUP_TOOLCHAIN=1.95.0 CARGO_BUILD_JOBS=6 RUST_TEST_THREADS=6
export TMPDIR=/home/zhuhe/code/litchi-worktrees/scratch/0767/tmp
export CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0767/dev
LOG=/home/zhuhe/code/litchi-worktrees/scratch/0767/gates.log
OUT=/home/zhuhe/code/litchi-worktrees/scratch/0767/gates-out
mkdir -p $OUT
: > $LOG
echo "worktree $(pwd) at $(git rev-parse HEAD); CARGO_TARGET_DIR=$CARGO_TARGET_DIR; CARGO_BUILD_JOBS=6; RUST_TEST_THREADS=6; TMPDIR=$TMPDIR; $(rustc --version)" >> $LOG
gate() {
  name=$1; shift
  echo "== $name: $*" >> $LOG
  "$@" > $OUT/$name.log 2>&1
  echo "exit=$?" >> $LOG
}
DEPS="-p litchi-cfb -p litchi-ole-common -p litchi-doc -p litchi-ppt -p litchi-xls -p litchi-vba -p litchi-ograph -p litchi-crypto -p litchi-sign"
OOXML="-p litchi-opc -p litchi-ooxml-common -p litchi-drawingml -p litchi-spreadsheet-drawing -p litchi-docx -p litchi-pptx -p litchi-xlsx -p litchi-xlsb"
gate fmt cargo fmt --all --check
gate check-cfb-dependents cargo check $DEPS --all-targets --locked --offline
gate check-ooxml-transitive cargo check $OOXML --all-targets --locked --offline
gate check-facade cargo check -p litchi --features doc,docx,ppt,pptx,xls,xlsx,xlsb,odt --all-targets --locked --offline
gate clippy-cfb-lib cargo clippy -p litchi-cfb --lib --no-deps --locked --offline -- -D warnings
gate clippy-cfb-all cargo clippy -p litchi-cfb --all-targets --no-deps --locked --offline -- -D warnings
gate doc-cfb env RUSTDOCFLAGS="-D warnings" cargo doc -p litchi-cfb --no-deps --locked --offline
gate boundaries python3 tools/check_crate_boundaries.py
gate non-iwork python3 tools/non_iwork_gate.py verify
gate perf-claims python3 tools/check_perf_claims.py --registry docs/performance/claim-registry-v1.json --repo-root . --mode structural
for crate in litchi-cfb litchi-ole-common litchi-doc litchi-ppt litchi-xls litchi-vba litchi-ograph litchi-crypto litchi-sign; do
  gate test-$crate cargo test -p $crate --locked --offline
done
gate test-facade cargo test -p litchi --features doc,docx,ppt,pptx,xls,xlsx,xlsb,odt --locked --offline
echo "done" >> $LOG
