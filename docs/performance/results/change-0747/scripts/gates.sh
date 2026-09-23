#!/bin/bash
# Change 0747 gates, run in the worktree at the committed candidate.
set -u
cd /home/zhuhe/code/litchi-worktrees/0747-xlsx-publication-audit-reuse
export CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0747 CARGO_BUILD_JOBS=6
LOG=/home/zhuhe/code/litchi-worktrees/scratch/0747/gates.txt
: > "$LOG"
mkdir -p /home/zhuhe/code/litchi-worktrees/scratch/0747/gate-logs
gate() {
  local name=$1; shift
  local full=/home/zhuhe/code/litchi-worktrees/scratch/0747/gate-logs/$name.log
  echo "=== $name" >> "$LOG"
  echo "\$ $*" >> "$LOG"
  "$@" > "$full" 2>&1
  local code=$?
  # Every test summary line, then every error or failure line, then the tail.
  grep -E "^test result" "$full" | sort | uniq -c >> "$LOG"
  local total
  total=$(grep -E "^test result" "$full" | sed -E 's/.* ([0-9]+) passed; ([0-9]+) failed; ([0-9]+) ignored.*/\1 \2 \3/' | awk '{p+=$1; f+=$2; i+=$3} END {if (NR) print "total: " p " passed, " f " failed, " i " ignored, in " NR " test binaries"}')
  [ -n "$total" ] && echo "$total" >> "$LOG"
  grep -E "^error|FAILED|panicked|^warning" "$full" | head -20 >> "$LOG"
  tail -3 "$full" >> "$LOG"
  echo "exit=$code" >> "$LOG"
  echo >> "$LOG"
}
echo "HEAD $(git rev-parse HEAD)" >> "$LOG"
gate fmt cargo fmt --all --check
gate check-touched-and-ooxml-dependents cargo check -p xml-minifier -p litchi-opc -p litchi-xlsx -p litchi-docx -p litchi-pptx -p litchi-xlsb -p litchi-ppt -p litchi-ooxml-common -p litchi-spreadsheet-drawing -p litchi --all-targets --locked --offline
gate check-odf-compile-only cargo check -p litchi-odf-common -p litchi-odt -p litchi-odp -p litchi-imgconv --all-targets --locked --offline
gate clippy-lib cargo clippy -p xml-minifier -p litchi-opc -p litchi-xlsx --lib --no-deps --locked --offline -- -D warnings
gate clippy-all-targets cargo clippy -p xml-minifier -p litchi-opc -p litchi-xlsx --all-targets --no-deps --locked --offline -- -D warnings
gate test-touched cargo test -p xml-minifier -p litchi-opc -p litchi-xlsx --locked --offline
gate test-ooxml-dependents cargo test -p litchi-docx -p litchi-pptx -p litchi-xlsb -p litchi-ooxml-common -p litchi-spreadsheet-drawing --locked --offline
gate test-facade cargo test -p litchi --features doc,docx,ppt,pptx,xls,xlsx,xlsb,odt --locked --offline
RUSTDOCFLAGS="-D warnings" gate doc cargo doc -p xml-minifier -p litchi-opc -p litchi-xlsx --no-deps --locked --offline
gate crate-boundaries python3 tools/check_crate_boundaries.py
gate non-iwork python3 tools/non_iwork_gate.py verify
gate perf-claims-structural python3 tools/check_perf_claims.py --registry docs/performance/claim-registry-v1.json --repo-root . --mode structural
echo "gates done" >> "$LOG"
