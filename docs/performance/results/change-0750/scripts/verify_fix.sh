#!/bin/bash
# Change 0750 follow-up: tests and gates on the fix, fresh target dir, TMPDIR on /home.
set -u
S=/home/zhuhe/code/litchi-worktrees/scratch/0750
cd /home/zhuhe/code/litchi-worktrees/0750-xml-audit-well-formedness-gaps
export TMPDIR=$S/tmp CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0750 CARGO_BUILD_JOBS=6
run() {
  local name=$1; shift
  "$@" > $S/logs/$name.log 2>&1
  echo "$name exit $?" >> $S/verify-status.txt
}
run xml-minifier cargo test -p xml-minifier --locked --offline --no-fail-fast
run ooxml cargo test -p litchi-opc -p litchi-ooxml-common -p litchi-docx -p litchi-xlsx -p litchi-pptx -p litchi-xlsb --locked --offline --no-fail-fast
run facade cargo test -p litchi --features doc,docx,ppt,pptx,xls,xlsx,xlsb,odt --locked --offline --no-fail-fast
run check-odf-ole cargo check -p litchi-odf-common -p litchi-odt -p litchi-ods -p litchi-odp -p litchi-xls -p litchi-ppt -p litchi-imgconv --all-targets --locked --offline
run clippy-lib cargo clippy -p xml-minifier -p litchi-opc -p litchi-docx -p litchi-xlsx --lib --no-deps --locked --offline -- -D warnings
run clippy-all cargo clippy -p xml-minifier -p litchi-opc -p litchi-docx -p litchi-xlsx --all-targets --no-deps --locked --offline -- -D warnings
RUSTDOCFLAGS="-D warnings" run doc cargo doc -p xml-minifier -p litchi-opc -p litchi-docx -p litchi-xlsx --no-deps --locked --offline
run fmt cargo fmt --all --check
echo done >> $S/verify-status.txt
