#!/bin/bash
# Change 0750 gates, in the worktree on the final candidate.
set -u
S=/home/zhuhe/code/litchi-worktrees/scratch/0750
G=$S/gates.txt
cd /home/zhuhe/code/litchi-worktrees/0750-xml-audit-well-formedness-gaps
export CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0750 CARGO_BUILD_JOBS=6
run() {
  echo "### $*" >> $G
  "$@" > $S/gate.log 2>&1
  local status=$?
  tail -4 $S/gate.log >> $G
  echo "exit $status" >> $G
  echo >> $G
}
: > $G
echo "Change 0750 gates, $(date -u +%Y-%m-%dT%H:%MZ), head $(git rev-parse --short HEAD) plus the working tree" >> $G
echo >> $G
run cargo fmt --all --check
run cargo check -p xml-minifier -p litchi-opc -p litchi-ooxml-common -p litchi-docx -p litchi-xlsx -p litchi-pptx -p litchi-xlsb -p litchi --all-targets --locked --offline
run cargo check -p litchi-odf-common -p litchi-odt -p litchi-ods -p litchi-odp -p litchi-xls -p litchi-ppt -p litchi-imgconv --all-targets --locked --offline
run cargo clippy -p xml-minifier -p litchi-opc -p litchi-docx -p litchi-xlsx --lib --no-deps --locked --offline -- -D warnings
run cargo clippy -p xml-minifier -p litchi-opc -p litchi-docx -p litchi-xlsx --all-targets --no-deps --locked --offline -- -D warnings
env RUSTDOCFLAGS="-D warnings" bash -c 'echo' > /dev/null
echo "### RUSTDOCFLAGS=\"-D warnings\" cargo doc -p xml-minifier -p litchi-opc -p litchi-docx -p litchi-xlsx --no-deps --locked --offline" >> $G
RUSTDOCFLAGS="-D warnings" cargo doc -p xml-minifier -p litchi-opc -p litchi-docx -p litchi-xlsx --no-deps --locked --offline > $S/gate.log 2>&1
status=$?; tail -3 $S/gate.log >> $G; echo "exit $status" >> $G; echo >> $G
run python3 tools/check_crate_boundaries.py
run python3 tools/non_iwork_gate.py verify
run python3 tools/check_perf_claims.py --registry docs/performance/claim-registry-v1.json --repo-root . --mode structural
echo "gates done" >> $S/gates-status.txt
