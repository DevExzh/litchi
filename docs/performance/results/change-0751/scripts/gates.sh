#!/usr/bin/env bash
# Run change 0751's gates from the worktree root and record each command's
# exit code and test counts. Usage: gates.sh OUT_DIR
set -u
out="$1"
mkdir -p "$out"
export CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0751
export CARGO_BUILD_JOBS=6
log="$out/gates.log"
: > "$log"

run() {
  local name="$1"
  shift
  "$@" > "$out/gate-$name.out" 2>&1
  local code=$?
  local counts
  counts=$(python3 - "$out/gate-$name.out" <<'PY'
import re, sys
text = open(sys.argv[1], encoding="utf-8", errors="replace").read()
passed = failed = ignored = 0
seen = False
for match in re.finditer(r"test result: \w+\. (\d+) passed; (\d+) failed; (\d+) ignored", text):
    seen = True
    passed += int(match.group(1)); failed += int(match.group(2)); ignored += int(match.group(3))
print(f"passed={passed} failed={failed} ignored={ignored}" if seen else "")
PY
)
  printf '%s\texit=%s\t%s\n' "$name" "$code" "$*" >> "$log"
  if [ -n "$counts" ]; then
    printf '    %s\n' "$counts" >> "$log"
  fi
}

run fmt cargo fmt --all --check
run check cargo check -p litchi-opc -p litchi-pptx -p litchi-ooxml-common -p litchi-docx -p litchi-xlsx -p litchi-xlsb -p litchi-spreadsheet-drawing -p litchi-ppt -p litchi --all-targets --locked --offline
run clippy-lib cargo clippy -p litchi-opc -p litchi-pptx --lib --no-deps --locked --offline -- -D warnings
run clippy-opc-all-targets cargo clippy -p litchi-opc --all-targets --no-deps --locked --offline -- -D warnings
run clippy-pptx-all-targets cargo clippy -p litchi-pptx --all-targets --no-deps --locked --offline --keep-going -- -D warnings
run test-core cargo test -p litchi-opc -p litchi-pptx --locked --offline
run test-pptx-release-cross-copy cargo test -p litchi-pptx --release --locked --offline --lib -- cross_copy_plan facade_memo
run test-dependents cargo test -p litchi-ooxml-common -p litchi-docx -p litchi-xlsx -p litchi-xlsb -p litchi-spreadsheet-drawing -p litchi-ppt --locked --offline
run test-facade cargo test -p litchi --features doc,docx,ppt,pptx,xls,xlsx,xlsb,odt --locked --offline
run doc env RUSTDOCFLAGS=-D\ warnings cargo doc -p litchi-opc -p litchi-pptx --no-deps --locked --offline
run crate-boundaries python3 tools/check_crate_boundaries.py
run non-iwork python3 tools/non_iwork_gate.py verify
run perf-claims python3 tools/check_perf_claims.py --registry docs/performance/claim-registry-v1.json --repo-root . --mode structural
echo "gates done" >> "$log"
