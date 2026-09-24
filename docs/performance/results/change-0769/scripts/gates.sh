#!/bin/bash
# Change 0769 gates, run from the worktree root in the foreground order below.
# Each command's exit code is appended to gates.log; a failure does not stop
# the later gates.
set -u
cd /home/zhuhe/code/litchi-worktrees/0769-cfb-mini-sector-open-read-agreement
source /home/zhuhe/code/litchi-worktrees/scratch/0769/env.sh
LOG=/home/zhuhe/code/litchi-worktrees/scratch/0769/gates.log
OUT=/home/zhuhe/code/litchi-worktrees/scratch/0769/gates-out
mkdir -p $OUT
: > $LOG
echo "worktree $(pwd) at $(git rev-parse HEAD); CARGO_TARGET_DIR=$CARGO_TARGET_DIR; CARGO_BUILD_JOBS=$CARGO_BUILD_JOBS; CARGO_INCREMENTAL=$CARGO_INCREMENTAL; RUST_TEST_THREADS=$RUST_TEST_THREADS; TMPDIR=$TMPDIR; $(rustc --version)" >> $LOG
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
gate doc-cfb-private env RUSTDOCFLAGS="-D warnings" cargo doc -p litchi-cfb --no-deps --document-private-items --locked --offline
for crate in litchi-cfb litchi-ole-common litchi-doc litchi-ppt litchi-xls litchi-vba litchi-ograph litchi-crypto litchi-sign; do
  gate test-$crate cargo test -p $crate --locked --offline
done
gate test-facade cargo test -p litchi --features doc,docx,ppt,pptx,xls,xlsx,xlsb,odt --locked --offline
echo "done" >> $LOG
