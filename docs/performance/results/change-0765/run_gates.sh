#!/usr/bin/env bash
# Change 0765 gates, run in the branch worktree with its own target directory.
set -u
cd /home/zhuhe/code/litchi-worktrees/0765-bom-prefixed-part-editing
export CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0765
export CARGO_BUILD_JOBS=6
export CARGO_INCREMENTAL=0
export TMPDIR=/home/zhuhe/code/litchi-worktrees/scratch/0765/tmpdir
LOG=/home/zhuhe/code/litchi-worktrees/scratch/0765/gates
: > "$LOG/gates.txt"
gate() { # name command...
  local name=$1; shift
  "$@" > "$LOG/$name.log" 2>&1
  local status=$?
  printf '%s\texit=%s\t%s\n' "$name" "$status" "$*" >> "$LOG/gates.txt"
}
TOUCHED="-p litchi-core -p litchi-opc -p litchi-ooxml-common -p litchi-drawingml -p litchi-xlsx -p litchi-pptx -p litchi-docx -p litchi-xlsb -p litchi-crypto -p litchi-formula -p litchi-xldm -p litchi-ole-common"
DEPENDENTS="-p litchi-spreadsheet-drawing -p litchi-doc -p litchi-xls -p litchi-ppt"
echo "HEAD $(git rev-parse HEAD)" >> "$LOG/gates.txt"
gate fmt cargo fmt --all --check
gate check-touched cargo check $TOUCHED $DEPENDENTS --all-targets --locked --offline
gate check-facade cargo check -p litchi --features doc,docx,ppt,pptx,xls,xlsx,xlsb,odt --all-targets --locked --offline
gate clippy-lib cargo clippy $TOUCHED --lib --no-deps --locked --offline -- -D warnings
for crate in litchi-core litchi-opc litchi-ooxml-common litchi-drawingml litchi-xlsx litchi-pptx litchi-docx litchi-xlsb litchi-crypto litchi-formula litchi-xldm litchi-ole-common; do
  gate "clippy-all-targets-$crate" cargo clippy -p "$crate" --all-targets --no-deps --locked --offline -- -D warnings
done
gate test-touched cargo test $TOUCHED $DEPENDENTS --locked --offline
gate test-facade cargo test -p litchi --features doc,docx,ppt,pptx,xls,xlsx,xlsb,odt --locked --offline
gate doc env RUSTDOCFLAGS="-D warnings" cargo doc $TOUCHED --no-deps --locked --offline
gate crate-boundaries python3 tools/check_crate_boundaries.py
gate non-iwork python3 tools/non_iwork_gate.py verify
gate perf-claims python3 tools/check_perf_claims.py --registry docs/performance/claim-registry-v1.json --repo-root . --mode structural
echo done >> "$LOG/gates.txt"
