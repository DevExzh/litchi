#!/usr/bin/env bash
# Every gate change 0632 ran, in order.  Run from the change's worktree.
#
#   gates.sh > gates.txt 2>&1
#
# `soapberry-zip` is the only crate whose production code changed, so the
# per-crate gates are its own plus the suite of every crate that depends on it.
# Change 0587 found two gates reachable from no per-crate gate: the harness's
# own suite and a feature-bearing `litchi` test.  Both are run.
set -uo pipefail
cd "$(dirname "${BASH_SOURCE[0]}")/../../../.." || exit 1
BUILD_CPUS=0-7,9-31   # CPU 8 is reserved for measured processes
run() {
  echo
  echo "================================================================"
  echo "### $*"
  echo "================================================================"
  taskset -c "$BUILD_CPUS" nice -n 10 "$@" 2>&1 | tail -25
  echo "--- exit ${PIPESTATUS[0]} ---"
}
run cargo fmt --all --check
run cargo clippy -p soapberry-zip --all-targets -j 6
run cargo doc -p soapberry-zip --no-deps -j 6
run cargo test -p soapberry-zip -j 6
run cargo test -p litchi-opc -j 6
run cargo test -p litchi-opc --all-features -j 6
run cargo test -p litchi-xlsx -p litchi-docx -p litchi-pptx -j 6
run cargo test -p litchi-ooxml-common -p litchi-core -j 6
run cargo test -p litchi --features docx,xlsx,pptx,xls -j 6
run cargo test -p litchi-odt -p litchi-odf-common -j 6
# Every remaining crate that depends on `soapberry-zip` directly, in one build.
run cargo test -p litchi-odc -p litchi-odg -p litchi-odp -p litchi-oth \
    -p litchi-odf-formula -p litchi-ppt -p litchi-sign -p litchi-iwa-archive -j 6
( cd tools/perf-baseline && run cargo test -j 6 )
echo
echo "all gates done"
