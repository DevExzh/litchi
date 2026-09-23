#!/usr/bin/env bash
# Gates for change 0748. Each command's exit code is recorded; the script never
# stops early so gates.txt lists every result.
set -u
W=/home/zhuhe/code/litchi-worktrees/0748-cfb-overlay-fingerprint-reuse
OUT=${1:-/home/zhuhe/code/litchi-worktrees/scratch/0748/gates.txt}
export CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0748
export CARGO_BUILD_JOBS=6
# Tests that stage atomic saves use std::env::temp_dir(); keep them off /tmp.
export TMPDIR=/home/zhuhe/code/litchi-worktrees/scratch/0748/tmp
cd "$W" || exit 1
: > "$OUT"
echo "HEAD $(git rev-parse HEAD)" >> "$OUT"
echo "rustc $(rustc --version)" >> "$OUT"
echo >> "$OUT"

gate() {
    echo "\$ $*" >> "$OUT"
    "$@" > /home/zhuhe/code/litchi-worktrees/scratch/0748/logs/gate.log 2>&1
    local code=$?
    grep -E "^test result|^error|^warning: unused|FAILED|panicked|Finished|could not compile|exit" \
        /home/zhuhe/code/litchi-worktrees/scratch/0748/logs/gate.log | tail -60 | sed 's/^/    /' >> "$OUT"
    echo "exit $code" >> "$OUT"
    echo >> "$OUT"
}

IN_SCOPE="-p litchi-cfb -p litchi-ole-common -p litchi-xls -p litchi-doc -p litchi-ppt -p litchi-vba -p litchi-sign -p litchi-crypto -p litchi-ograph -p litchi-docx -p litchi-pptx -p litchi-xlsx -p litchi-xlsb -p litchi-opc -p litchi-drawingml -p litchi-ooxml-common -p litchi-spreadsheet-drawing"

gate cargo fmt --all --check
gate cargo fmt --manifest-path tools/perf-baseline/Cargo.toml --check
gate cargo check $IN_SCOPE --all-targets --locked --offline
gate cargo check -p litchi --features doc,docx,ppt,pptx,xls,xlsx,xlsb,odt --all-targets --locked --offline
gate cargo clippy -p litchi-cfb -p litchi-ole-common -p litchi-xls --lib --no-deps --locked --offline -- -D warnings
gate cargo clippy -p litchi-cfb -p litchi-ole-common --all-targets --no-deps --locked --offline -- -D warnings
gate cargo clippy -p litchi-xls --all-targets --no-deps --locked --offline -- -D warnings
gate cargo test -p litchi-cfb --locked --offline
gate cargo test -p litchi-ole-common --locked --offline
gate cargo test -p litchi-xls --locked --offline
gate cargo test -p litchi-doc --locked --offline
gate cargo test -p litchi-ppt --locked --offline
gate cargo test -p litchi-vba -p litchi-sign -p litchi-crypto -p litchi-ograph --locked --offline
gate cargo test -p litchi --features doc,docx,ppt,pptx,xls,xlsx,xlsb,odt --locked --offline
RUSTDOCFLAGS="-D warnings" gate cargo doc -p litchi-cfb -p litchi-ole-common -p litchi-xls --no-deps --locked --offline
gate cargo test --manifest-path tools/perf-baseline/Cargo.toml --locked --offline
gate python3 -m unittest tools.test_perf_abba_summary
gate python3 tools/validate_crud_coverage_index.py --index docs/performance/crud-coverage-index-v1.json --catalog docs/performance/results/perf-corpus-manifest-v2.json --selector-source tools/perf-baseline/src/lib.rs --checklist docs/CRUD_Scenario_Checklist.md --repo-root .
gate python3 tools/check_crate_boundaries.py
gate python3 tools/non_iwork_gate.py verify
gate python3 tools/check_perf_claims.py --registry docs/performance/claim-registry-v1.json --repo-root . --mode structural
echo "done" >> "$OUT"
