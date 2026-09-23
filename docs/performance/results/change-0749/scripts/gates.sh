#!/bin/bash
# Change 0749 gates. Run from the worktree root after measurement has finished.
# Every command's exit code is recorded; test totals are summed per command.
set -u
cd /home/zhuhe/code/litchi-worktrees/0749-cfb-reuse-plan-validation
export RUSTUP_TOOLCHAIN=1.95.0
export CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0749/ws
export CARGO_BUILD_JOBS=6
export RUST_TEST_THREADS=6
OUT=/home/zhuhe/code/litchi-worktrees/scratch/0749/gates
mkdir -p "$OUT"
LOG="$OUT/gates.txt"
DEPENDENTS="-p litchi-cfb -p litchi-ole-common -p litchi-doc -p litchi-ppt -p litchi-xls -p litchi-vba -p litchi-ograph -p litchi-crypto -p litchi-sign"
FACADE_FEATURES="doc,docx,ppt,pptx,xls,xlsx,xlsb,odt"

{
  echo "# Change 0749 gates"
  echo
  echo "Worktree $(pwd) at $(git rev-parse HEAD); CARGO_TARGET_DIR=$CARGO_TARGET_DIR; CARGO_BUILD_JOBS=6; RUST_TEST_THREADS=6; $(rustc --version)."
  echo "Each command ran from the repository root; exit codes as recorded."
  echo
} > "$LOG"

run() {
  local name="$1"
  shift
  echo "== $name: $*" >> "$LOG"
  "$@" > "$OUT/$name.log" 2>&1
  echo "exit=$?" >> "$LOG"
}

run fmt cargo fmt --all --check
run check-cfb-dependents cargo check $DEPENDENTS --all-targets --locked --offline
run check-facade cargo check -p litchi --features "$FACADE_FEATURES" --all-targets --locked --offline
run clippy-cfb-lib cargo clippy -p litchi-cfb --lib --no-deps --locked --offline -- -D warnings
run clippy-cfb-all cargo clippy -p litchi-cfb --all-targets --no-deps --locked --offline -- -D warnings
run doc-cfb env RUSTDOCFLAGS="-D warnings" cargo doc -p litchi-cfb --no-deps --locked --offline
run boundaries python3 tools/check_crate_boundaries.py
run non-iwork python3 tools/non_iwork_gate.py verify
run perf-claims python3 tools/check_perf_claims.py --registry docs/performance/claim-registry-v1.json --repo-root . --mode structural
for crate in litchi-cfb litchi-ole-common litchi-doc litchi-ppt litchi-xls litchi-vba litchi-ograph litchi-crypto litchi-sign; do
  run "test-$crate" cargo test -p "$crate" --locked --offline
done
run test-facade cargo test -p litchi --features "$FACADE_FEATURES" --locked --offline

{
  echo
  echo "Test totals (sum of all test binaries and doctests):"
  for log in "$OUT"/test-*.log; do
    name=$(basename "$log" .log)
    python3 - "$log" "$name" <<'EOF'
import re, sys
text = open(sys.argv[1]).read()
passed = failed = ignored = 0
for match in re.finditer(r"test result: \w+\. (\d+) passed; (\d+) failed; (\d+) ignored", text):
    passed += int(match.group(1)); failed += int(match.group(2)); ignored += int(match.group(3))
print(f"  {sys.argv[2]}: passed={passed} failed={failed} ignored={ignored}")
EOF
  done
  echo "  boundaries: $(head -1 "$OUT/boundaries.log")"
  echo "  non-iwork: $(tail -1 "$OUT/non-iwork.log")"
  echo "  perf-claims: $(tail -1 "$OUT/perf-claims.log")"
} >> "$LOG"
echo done >> "$OUT/finished"
