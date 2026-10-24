#!/usr/bin/env bash
# HISTORICAL/SUPERSEDED: this c796 recipe is retained only for provenance.
# Use 472b-qualification-commands.sh for the current candidate; do not use or
# relabel c796 receipts, binaries or source archives as 472b.
# Reproducible W8 c796 qualification recipe.  All outputs, target trees,
# binaries and receipts stay outside the repository.  This file only runs a
# selected stage; it never chooses a source overlay or a baseline implicitly.
set -Eeuo pipefail

SCRIPT_DIR=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
ROOT=${PROTO_ROOT:-$(git -C "$SCRIPT_DIR" rev-parse --show-toplevel)}
W8_PREP=${W8_PREP:-}
W8_BASELINE_ROOT=${W8_BASELINE_ROOT:-}
W8_BASELINE_ARCHIVE=${W8_BASELINE_ARCHIVE:-}
CANDIDATE_SOURCE=${CANDIDATE_SOURCE:-$ROOT}
CARGO=${CARGO:-cargo}
TARGET=x86_64-unknown-linux-gnu
PROFILE=release
FINAL_COMMIT=c796d0acc24c7c6e6fd08980b9e31af6bb86b18b
BASELINE_COMMIT=f2bbbc8d07ab5da26c180f1ff74813d39f632d54
PROBE_REL=crates/litchi-ops/examples/file_w8_qualified_probe.rs
BASE_ADAPTER_REL=crates/litchi-ops/examples/file_w8_original_baseline_probe.rs
PROVENANCE="$ROOT/tools/w8_file_provenance.py"
BASELINE_TOOL="$ROOT/tools/w8_original_baseline.py"
POLICY="$ROOT/docs/unified-ops/results/0362yh/original-to-final-policy.json"
RUNNER="$ROOT/docs/unified-ops/results/0362h/abba.sh"
ANALYZER="$ROOT/docs/unified-ops/results/0362h/analyze.py"
VALIDATOR="$ROOT/docs/unified-ops/results/0362h/validate.py"

BASE_OUT="$W8_PREP/baseline/provenance"
FINAL_OUT="$W8_PREP/candidate"
BASE_TARGET="$W8_PREP/target-before"
FINAL_TARGET="$W8_PREP/target-after"
BASE_BIN="$BASE_TARGET/$TARGET/$PROFILE/examples/file_w8_original_baseline_probe"
FINAL_BIN="$FINAL_TARGET/$TARGET/$PROFILE/examples/file_w8_qualified_probe"
BASE_FIXTURE="$BASE_OUT/litchi_w8_fixture_descriptor.json"
FINAL_FIXTURE="$FINAL_OUT/fixture.json"
BASE_PROJECTION="$BASE_OUT/litchi_w8_dependency_projection.json"
FINAL_PROJECTION="$FINAL_OUT/projection.json"
BASE_ASSEMBLY="$BASE_OUT/original-baseline-source-assembly.json"
FINAL_ASSEMBLY="$FINAL_OUT/source-assembly.json"
BASE_MANIFEST="$W8_PREP/endpoint-before.json"
FINAL_MANIFEST="$W8_PREP/endpoint-after.json"
PAIR_RECEIPT="$W8_PREP/original-to-final-pair.json"
ABBA_OUT="$W8_PREP/abba-c796"

require_inputs() {
  : "${W8_PREP:?set W8_PREP to an external scratch directory}"
  : "${W8_BASELINE_ROOT:?set W8_BASELINE_ROOT to the immutable f2 workspace}"
  : "${W8_BASELINE_ARCHIVE:?set W8_BASELINE_ARCHIVE to the immutable f2 source archive}"
}

prepare() {
  require_inputs
  mkdir -p "$BASE_OUT" "$FINAL_OUT"
  python3 "$BASELINE_TOOL" prepare \
    --git-root "$CANDIDATE_SOURCE" \
    --workspace "$W8_BASELINE_ROOT" \
    --probe "$W8_BASELINE_ROOT/$PROBE_REL" \
    --lock "$W8_BASELINE_ROOT/Cargo.lock" \
    --lock-source-commit "$BASELINE_COMMIT" \
    --lock-origin 'historical f2 dependency resolution; immutable 139441-byte lock' \
    --source-archive "$W8_BASELINE_ARCHIVE" \
    --output-dir "$BASE_OUT" --target "$TARGET" --profile "$PROFILE"

  test ! -e "$FINAL_OUT/c796-source.tar"
  git -C "$CANDIDATE_SOURCE" archive --format=tar \
    --output="$FINAL_OUT/c796-source.tar" "$FINAL_COMMIT"
  python3 "$PROVENANCE" fixture \
    --probe "$CANDIDATE_SOURCE/$PROBE_REL" --output "$FINAL_FIXTURE"
  python3 "$PROVENANCE" dependency \
    --workspace "$CANDIDATE_SOURCE" \
    --package-manifest "$CANDIDATE_SOURCE/crates/litchi-ops/Cargo.toml" \
    --lock "$CANDIDATE_SOURCE/Cargo.lock" \
    --toolchain "$CANDIDATE_SOURCE/rust-toolchain.toml" \
    --probe "$CANDIDATE_SOURCE/$PROBE_REL" --fixture "$FINAL_FIXTURE" \
    --all-features --target "$TARGET" --profile "$PROFILE" \
    --output "$FINAL_PROJECTION"
  python3 "$PROVENANCE" source \
    --source "$FINAL_OUT/c796-source.tar" --source-commit "$FINAL_COMMIT" \
    --overlay-kind none \
    --description 'stable c796 production source; no production compatibility overlay' \
    --output "$FINAL_ASSEMBLY"
}

build() {
  require_inputs
  (cd "$W8_BASELINE_ROOT" && \
    CARGO_BUILD_JOBS=1 CARGO_TARGET_DIR="$BASE_TARGET" "$CARGO" \
      build --offline --locked -p litchi-ops --all-features \
      --example file_w8_original_baseline_probe)
  (cd "$CANDIDATE_SOURCE" && \
    CARGO_BUILD_JOBS=1 CARGO_TARGET_DIR="$FINAL_TARGET" "$CARGO" \
      build --offline --locked -p litchi-ops --all-features \
      --example file_w8_qualified_probe)
}

validate() {
  require_inputs
  # The before binary compiles the generated consumer adapter.  The corrected
  # probe remains the shared workload identity in the endpoint; the source
  # assembly and adapter patch bind the compiled adapter to that workload.
  BASE_BUILD_COMMAND="CARGO_BUILD_JOBS=1 CARGO_TARGET_DIR=$BASE_TARGET $CARGO build --offline --locked -p litchi-ops --all-features --example file_w8_original_baseline_probe"
  FINAL_BUILD_COMMAND="CARGO_BUILD_JOBS=1 CARGO_TARGET_DIR=$FINAL_TARGET $CARGO build --offline --locked -p litchi-ops --all-features --example file_w8_qualified_probe"
  python3 -B "$VALIDATOR" \
    --source "$W8_BASELINE_ARCHIVE" --probe "$W8_BASELINE_ROOT/$PROBE_REL" \
    --lock "$W8_BASELINE_ROOT/Cargo.lock" --source-commit "$BASELINE_COMMIT" \
    --build-command "$BASE_BUILD_COMMAND" "$BASE_BIN" \
    "$W8_PREP/validation-before.json"
  python3 -B "$VALIDATOR" \
    --source "$FINAL_OUT/c796-source.tar" --probe "$CANDIDATE_SOURCE/$PROBE_REL" \
    --lock "$CANDIDATE_SOURCE/Cargo.lock" --source-commit "$FINAL_COMMIT" \
    --build-command "$FINAL_BUILD_COMMAND" "$FINAL_BIN" \
    "$W8_PREP/validation-after.json"

  python3 "$PROVENANCE" endpoint --binary "$BASE_BIN" \
    --source "$W8_BASELINE_ARCHIVE" --lock "$W8_BASELINE_ROOT/Cargo.lock" \
    --probe "$W8_BASELINE_ROOT/$PROBE_REL" --fixture "$BASE_FIXTURE" \
    --dependency-projection "$BASE_PROJECTION" --workspace "$W8_BASELINE_ROOT" \
    --workspace-manifest "$W8_BASELINE_ROOT/Cargo.toml" \
    --package-manifest "$W8_BASELINE_ROOT/crates/litchi-ops/Cargo.toml" \
    --toolchain "$W8_BASELINE_ROOT/rust-toolchain.toml" \
    --source-assembly "$BASE_ASSEMBLY" \
    --build-receipt "$W8_PREP/build-before.json" \
    --validation-receipt "$W8_PREP/validation-before.json" \
    --output "$BASE_MANIFEST" --source-commit "$BASELINE_COMMIT" \
    --all-features --target "$TARGET" --profile "$PROFILE"
  python3 "$PROVENANCE" endpoint --binary "$FINAL_BIN" \
    --source "$FINAL_OUT/c796-source.tar" --lock "$CANDIDATE_SOURCE/Cargo.lock" \
    --probe "$CANDIDATE_SOURCE/$PROBE_REL" --fixture "$FINAL_FIXTURE" \
    --dependency-projection "$FINAL_PROJECTION" --workspace "$CANDIDATE_SOURCE" \
    --workspace-manifest "$CANDIDATE_SOURCE/Cargo.toml" \
    --package-manifest "$CANDIDATE_SOURCE/crates/litchi-ops/Cargo.toml" \
    --toolchain "$CANDIDATE_SOURCE/rust-toolchain.toml" \
    --source-assembly "$FINAL_ASSEMBLY" \
    --build-receipt "$W8_PREP/build-after.json" \
    --validation-receipt "$W8_PREP/validation-after.json" \
    --output "$FINAL_MANIFEST" --source-commit "$FINAL_COMMIT" \
    --all-features --target "$TARGET" --profile "$PROFILE"
}

pair() {
  require_inputs
  python3 "$ROOT/tools/w8_projection_policy.py" pair \
    --policy "$POLICY" --before-manifest "$BASE_MANIFEST" \
    --after-manifest "$FINAL_MANIFEST" --output "$PAIR_RECEIPT"
}

abba() {
  require_inputs
  CPU=${CPU:-$(python3 -c 'import os; print(min(os.sched_getaffinity(0)))')}
  ROUNDS=${ROUNDS:-5}
  bash "$RUNNER" "$BASE_BIN" "$FINAL_BIN" "$ABBA_OUT" "$ROUNDS" "$CPU" \
    --before-manifest "$BASE_MANIFEST" --after-manifest "$FINAL_MANIFEST" \
    --original-to-final-policy "$POLICY" \
    --original-to-final-receipt "$PAIR_RECEIPT"
  python3 -B "$ANALYZER" "$ABBA_OUT"
}

allocator() {
  echo 'W8 File has no allocator-instrumented probe or allocator ABBA protocol; no allocator claim is emitted.' >&2
  return 2
}

case "${1:-}" in
  prepare|build|validate|pair|abba|allocator) "$1" ;;
  *) echo "usage: $0 {prepare|build|validate|pair|abba|allocator}" >&2; exit 2 ;;
esac
