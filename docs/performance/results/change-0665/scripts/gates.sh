#!/bin/bash
# Change 0665 gate sequence. usage: gates.sh <worktree> <outdir>
set -u
W="$1"; OUT="$2"; mkdir -p "$OUT"
cd "$W" || exit 1
export CARGO_BUILD_JOBS="${CARGO_BUILD_JOBS:-2}"
run() { # run <label> <command...>
  local label="$1"; shift
  echo "=== $label" >> "$OUT/gates.txt"
  echo "\$ $*" >> "$OUT/gates.txt"
  "$@" > "$OUT/$label.out" 2>&1
  local rc=$?
  tail -25 "$OUT/$label.out" >> "$OUT/gates.txt"
  echo "[exit $rc]" >> "$OUT/gates.txt"
  echo >> "$OUT/gates.txt"
  echo "$label rc=$rc"
}
: > "$OUT/gates.txt"
run fmt cargo fmt --all --check
run clippy cargo clippy -p litchi-opc -p litchi-docx --all-targets
run test-opc-docx cargo test -p litchi-opc -p litchi-docx
run doc cargo doc -p litchi-opc -p litchi-docx --no-deps
run test-consumers cargo test -p litchi-xlsx -p litchi-pptx
run test-facade cargo test -p litchi --features docx,xlsx,pptx,xls
run test-perf-baseline cargo test --manifest-path tools/perf-baseline/Cargo.toml
run non-iwork python3 tools/non_iwork_gate.py verify
