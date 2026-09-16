#!/usr/bin/env bash
# Every gate change 0664 ran, in order, with each tail appended to gates.txt.
# Usage: run_gates.sh <worktree> <gates.txt>
set -uo pipefail

WORKTREE="${1:?worktree}"
OUT="${2:?gates.txt}"
: > "$OUT"

section() { printf '\n========== %s ==========\n' "$1" >> "$OUT"; }

run() { # <label> <cwd> <command...>
  local label="$1" cwd="$2"; shift 2
  section "$label"
  ( cd "$cwd" && "$@" ) >> "$OUT" 2>&1
  printf '[exit %s] %s\n' "$?" "$label" >> "$OUT"
}

run "cargo fmt --all --check (tools/perf-baseline)" "$WORKTREE/tools/perf-baseline" \
  cargo fmt --all --check
run "cargo clippy --release --locked --all-targets (tools/perf-baseline)" "$WORKTREE/tools/perf-baseline" \
  cargo clippy --release --locked --all-targets
run "cargo clippy --release --locked --all-targets --features allocator-metrics" "$WORKTREE/tools/perf-baseline" \
  cargo clippy --release --locked --all-targets --features allocator-metrics
run "cargo doc --no-deps (tools/perf-baseline)" "$WORKTREE/tools/perf-baseline" \
  cargo doc --no-deps
run "cargo test --release --locked (tools/perf-baseline)" "$WORKTREE/tools/perf-baseline" \
  cargo test --release --locked
run "python3 tools/non_iwork_gate.py harness-tests" "$WORKTREE" \
  python3 tools/non_iwork_gate.py harness-tests
run "python3 tools/validate_crud_coverage_index.py" "$WORKTREE" \
  python3 tools/validate_crud_coverage_index.py
run "python3 -m unittest tools.test_corpus_manifest_v2 tools.test_crud_coverage_index tools.test_perf_baseline_source_policy tools.test_perf_corpus_binding" "$WORKTREE" \
  python3 -m unittest tools.test_corpus_manifest_v2 tools.test_crud_coverage_index tools.test_perf_baseline_source_policy tools.test_perf_corpus_binding
run "python3 tools/non_iwork_gate.py verify" "$WORKTREE" \
  python3 tools/non_iwork_gate.py verify
run "derive_marker_shape.py verify" "$WORKTREE" \
  python3 docs/performance/results/change-0664/scripts/derive_marker_shape.py verify

echo "gates done"
