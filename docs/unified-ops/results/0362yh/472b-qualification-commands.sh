#!/usr/bin/env bash
# Reproducible W8 472b qualification recipe.  All outputs, target trees,
# binaries and receipts stay outside the repository.  This file only runs a
# selected stage; it never chooses a source overlay or a baseline implicitly.
set -Eeuo pipefail

SCRIPT_DIR=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
ROOT=${PROTO_ROOT:-$(git -C "$SCRIPT_DIR" rev-parse --show-toplevel)}
W8_PREP=${W8_PREP:-}
W8_BASELINE_ROOT=${W8_BASELINE_ROOT:-}
W8_BASELINE_ARCHIVE=${W8_BASELINE_ARCHIVE:-}
CANDIDATE_SOURCE=${CANDIDATE_SOURCE:-}
CARGO=${CARGO:-cargo}
CARGO_BIN=
RUSTC_BIN=
TARGET=x86_64-unknown-linux-gnu
PROFILE=release
FINAL_COMMIT=472b6705475b073e7fe08f6e11ee17fb1e29f838
BASELINE_COMMIT=f2bbbc8d07ab5da26c180f1ff74813d39f632d54
CANDIDATE_LOCK_BYTES=139606
CANDIDATE_LOCK_SHA256=4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1
CANDIDATE_WORKSPACE_BYTES=10672
CANDIDATE_WORKSPACE_SHA256=f5ec007892023b38f1deaef1ce20714db58b58448a3bea088be3b847b84570de
CANDIDATE_PACKAGE_BYTES=13718
CANDIDATE_PACKAGE_SHA256=004d3b3cba8c035d209665a76c9cfe3932f12a9c198a4aea9fa314ef8ebb3a39
CANDIDATE_TOOLCHAIN_BYTES=86
CANDIDATE_TOOLCHAIN_SHA256=e3a213e0d222e94d213cafbc20932eb3f76c643b4dd63756acf95192df2aa310
CANDIDATE_PROBE_BYTES=75035
CANDIDATE_PROBE_SHA256=21c8e3833c96411a612a2631f98093bd70cf0a5c28b2952c3d15873f464e54b0
CANDIDATE_SOURCE_BYTES=6634147840
CANDIDATE_SOURCE_SHA256=44c11e933974c8fa3d6a67b46739bedeef866e60fbd4360ffe7f3bb8a4dc2274
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
ABBA_OUT="$W8_PREP/abba-472b"
# validate.py writes a receipt directory containing the canonical receipt and
# the route artifact.  Endpoint provenance binds the canonical receipt file,
# not the enclosing directory.
BASE_VALIDATION_RECEIPT="$W8_PREP/validation-before.json/receipt.json"
FINAL_VALIDATION_RECEIPT="$W8_PREP/validation-after.json/receipt.json"

require_inputs() {
  : "${W8_PREP:?set W8_PREP to an external scratch directory}"
  : "${W8_BASELINE_ROOT:?set W8_BASELINE_ROOT to the immutable f2 workspace}"
  : "${W8_BASELINE_ARCHIVE:?set W8_BASELINE_ARCHIVE to the immutable f2 source archive}"
  : "${CANDIDATE_SOURCE:?set CANDIDATE_SOURCE to the absolute clean 472b checkout}"
  case "$CANDIDATE_SOURCE" in
    /*) ;;
    *) echo "CANDIDATE_SOURCE must be an absolute path" >&2; return 2 ;;
  esac
  CANDIDATE_SOURCE=$(realpath -- "$CANDIDATE_SOURCE")
  test "$BASE_TARGET" != "$FINAL_TARGET" || { echo "before/after target directories must differ" >&2; return 2; }
  verify_candidate_source
}

verify_artifact_identity() {
  local path=$1 expected_bytes=$2 expected_sha256=$3 actual_bytes actual_sha256
  test -f "$path" && ! test -L "$path" || { echo "artifact is not a regular file: $path" >&2; return 2; }
  actual_bytes=$(stat -c '%s' "$path")
  test "$actual_bytes" = "$expected_bytes" || {
    echo "artifact byte identity differs: $path ($actual_bytes != $expected_bytes)" >&2
    return 2
  }
  actual_sha256=$(sha256sum "$path" | awk '{print $1}')
  test "$actual_sha256" = "$expected_sha256" || {
    echo "artifact hash identity differs: $path ($actual_sha256 != $expected_sha256)" >&2
    return 2
  }
}

verify_candidate_source() {
  case "$CANDIDATE_SOURCE" in
    /*) ;;
    *) echo "CANDIDATE_SOURCE must be an absolute path" >&2; return 2 ;;
  esac
  test -d "$CANDIDATE_SOURCE" || { echo "candidate source is not a directory: $CANDIDATE_SOURCE" >&2; return 2; }
  local git_root head status
  git_root=$(git -C "$CANDIDATE_SOURCE" rev-parse --show-toplevel 2>/dev/null) || {
    echo "candidate source is not a git checkout: $CANDIDATE_SOURCE" >&2
    return 2
  }
  git_root=$(realpath -- "$git_root")
  test "$git_root" = "$CANDIDATE_SOURCE" || {
    echo "candidate source git root differs from requested checkout" >&2
    return 2
  }
  head=$(git -C "$CANDIDATE_SOURCE" rev-parse --verify HEAD^{commit} 2>/dev/null) || {
    echo "candidate source HEAD is unavailable" >&2
    return 2
  }
  test "$head" = "$FINAL_COMMIT" || {
    echo "candidate source HEAD is not the reviewed 472b commit: $head" >&2
    return 2
  }
  status=$(git -C "$CANDIDATE_SOURCE" status --porcelain=v1 --untracked-files=all) || return 2
  test -z "$status" || {
    echo "candidate source checkout is dirty" >&2
    return 2
  }
  verify_artifact_identity "$CANDIDATE_SOURCE/Cargo.lock" "$CANDIDATE_LOCK_BYTES" "$CANDIDATE_LOCK_SHA256"
  verify_artifact_identity "$CANDIDATE_SOURCE/Cargo.toml" "$CANDIDATE_WORKSPACE_BYTES" "$CANDIDATE_WORKSPACE_SHA256"
  verify_artifact_identity "$CANDIDATE_SOURCE/crates/litchi-ops/Cargo.toml" "$CANDIDATE_PACKAGE_BYTES" "$CANDIDATE_PACKAGE_SHA256"
  verify_artifact_identity "$CANDIDATE_SOURCE/rust-toolchain.toml" "$CANDIDATE_TOOLCHAIN_BYTES" "$CANDIDATE_TOOLCHAIN_SHA256"
  verify_artifact_identity "$CANDIDATE_SOURCE/$PROBE_REL" "$CANDIDATE_PROBE_BYTES" "$CANDIDATE_PROBE_SHA256"
}

verify_candidate_archive() {
  verify_artifact_identity "$1" "$CANDIDATE_SOURCE_BYTES" "$CANDIDATE_SOURCE_SHA256"
}

ensure_empty_target() {
  local target=$1
  if test -e "$target"; then
    test -d "$target" || { echo "target path is not a directory: $target" >&2; return 2; }
    if find "$target" -mindepth 1 -print -quit | grep -q .; then
      echo "target directory must be empty before a provenance-bound build: $target" >&2
      return 2
    fi
  fi
}

resolve_tools() {
  CARGO_BIN=$(command -v "$CARGO" || true)
  RUSTC_BIN=$(command -v rustc || true)
  test -n "$CARGO_BIN" || { echo "cargo shim is not on PATH: $CARGO" >&2; return 2; }
  test -n "$RUSTC_BIN" || { echo "rustc is not on PATH" >&2; return 2; }
  CARGO_BIN=$(realpath -- "$CARGO_BIN")
  RUSTC_BIN=$(realpath -- "$RUSTC_BIN")
  test -x "$CARGO_BIN" || { echo "cargo shim is not executable: $CARGO_BIN" >&2; return 2; }
  test -x "$RUSTC_BIN" || { echo "rustc is not executable: $RUSTC_BIN" >&2; return 2; }
}

build_commands() {
  resolve_tools
  BASE_BUILD_COMMAND="CARGO_BUILD_JOBS=1 CARGO_TARGET_DIR=$BASE_TARGET $CARGO_BIN build --offline --locked --target $TARGET --release -p litchi-ops --all-features --example file_w8_original_baseline_probe"
  FINAL_BUILD_COMMAND="CARGO_BUILD_JOBS=1 CARGO_TARGET_DIR=$FINAL_TARGET $CARGO_BIN build --offline --locked --target $TARGET --release -p litchi-ops --all-features --example file_w8_qualified_probe"
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

  test ! -e "$FINAL_OUT/472b-source.tar" && test ! -L "$FINAL_OUT/472b-source.tar" || {
    echo "refusing to overwrite an existing candidate source archive" >&2
    return 2
  }
  git -C "$CANDIDATE_SOURCE" archive --format=tar \
    --output="$FINAL_OUT/472b-source.tar" "$FINAL_COMMIT"
  verify_candidate_archive "$FINAL_OUT/472b-source.tar"
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
    --source "$FINAL_OUT/472b-source.tar" --source-commit "$FINAL_COMMIT" \
    --overlay-kind none \
    --description 'stable 472b production source; no production compatibility overlay' \
    --output "$FINAL_ASSEMBLY"
}

build() {
  require_inputs
  verify_candidate_archive "$FINAL_OUT/472b-source.tar"
  ensure_empty_target "$BASE_TARGET"
  ensure_empty_target "$FINAL_TARGET"
  build_commands
  (cd "$W8_BASELINE_ROOT" && \
    CARGO_BUILD_JOBS=1 CARGO_TARGET_DIR="$BASE_TARGET" "$CARGO_BIN" \
      build --offline --locked --target "$TARGET" --release -p litchi-ops --all-features \
      --example file_w8_original_baseline_probe)
  python3 "$PROVENANCE" build-receipt \
    --source "$W8_BASELINE_ARCHIVE" --source-assembly "$BASE_ASSEMBLY" \
    --binary "$BASE_BIN" --lock "$W8_BASELINE_ROOT/Cargo.lock" \
    --probe "$W8_BASELINE_ROOT/$PROBE_REL" --fixture "$BASE_FIXTURE" \
    --dependency-projection "$BASE_PROJECTION" --workspace "$W8_BASELINE_ROOT" \
    --workspace-manifest "$W8_BASELINE_ROOT/Cargo.toml" \
    --package-manifest "$W8_BASELINE_ROOT/crates/litchi-ops/Cargo.toml" \
    --toolchain "$W8_BASELINE_ROOT/rust-toolchain.toml" --cargo "$CARGO_BIN" \
    --compiler-artifact "rustc=$RUSTC_BIN" --source-commit "$BASELINE_COMMIT" \
    --build-command "$BASE_BUILD_COMMAND" --all-features --target "$TARGET" \
    --profile "$PROFILE" --output "$W8_PREP/build-before.json"
  (cd "$CANDIDATE_SOURCE" && \
    CARGO_BUILD_JOBS=1 CARGO_TARGET_DIR="$FINAL_TARGET" "$CARGO_BIN" \
      build --offline --locked --target "$TARGET" --release -p litchi-ops --all-features \
      --example file_w8_qualified_probe)
  python3 "$PROVENANCE" build-receipt \
    --source "$FINAL_OUT/472b-source.tar" --source-assembly "$FINAL_ASSEMBLY" \
    --binary "$FINAL_BIN" --lock "$CANDIDATE_SOURCE/Cargo.lock" \
    --probe "$CANDIDATE_SOURCE/$PROBE_REL" --fixture "$FINAL_FIXTURE" \
    --dependency-projection "$FINAL_PROJECTION" --workspace "$CANDIDATE_SOURCE" \
    --workspace-manifest "$CANDIDATE_SOURCE/Cargo.toml" \
    --package-manifest "$CANDIDATE_SOURCE/crates/litchi-ops/Cargo.toml" \
    --toolchain "$CANDIDATE_SOURCE/rust-toolchain.toml" --cargo "$CARGO_BIN" \
    --compiler-artifact "rustc=$RUSTC_BIN" --source-commit "$FINAL_COMMIT" \
    --build-command "$FINAL_BUILD_COMMAND" --all-features --target "$TARGET" \
    --profile "$PROFILE" --output "$W8_PREP/build-after.json"
}

validate() {
  require_inputs
  verify_candidate_archive "$FINAL_OUT/472b-source.tar"
  build_commands
  test -s "$W8_PREP/build-before.json" || { echo "missing complete before build receipt; run build first" >&2; return 2; }
  test -s "$W8_PREP/build-after.json" || { echo "missing complete after build receipt; run build first" >&2; return 2; }
  # The before binary compiles the generated consumer adapter.  The corrected
  # probe remains the shared workload identity in the endpoint; the source
  # assembly and adapter patch bind the compiled adapter to that workload.
  python3 -B "$VALIDATOR" \
    --source "$W8_BASELINE_ARCHIVE" --probe "$W8_BASELINE_ROOT/$PROBE_REL" \
    --lock "$W8_BASELINE_ROOT/Cargo.lock" --source-commit "$BASELINE_COMMIT" \
    --build-command "$BASE_BUILD_COMMAND" "$BASE_BIN" \
    "$W8_PREP/validation-before.json"
  python3 -B "$VALIDATOR" \
    --source "$FINAL_OUT/472b-source.tar" --probe "$CANDIDATE_SOURCE/$PROBE_REL" \
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
    --validation-receipt "$BASE_VALIDATION_RECEIPT" \
    --output "$BASE_MANIFEST" --source-commit "$BASELINE_COMMIT" \
    --all-features --target "$TARGET" --profile "$PROFILE"
  python3 "$PROVENANCE" endpoint --binary "$FINAL_BIN" \
    --source "$FINAL_OUT/472b-source.tar" --lock "$CANDIDATE_SOURCE/Cargo.lock" \
    --probe "$CANDIDATE_SOURCE/$PROBE_REL" --fixture "$FINAL_FIXTURE" \
    --dependency-projection "$FINAL_PROJECTION" --workspace "$CANDIDATE_SOURCE" \
    --workspace-manifest "$CANDIDATE_SOURCE/Cargo.toml" \
    --package-manifest "$CANDIDATE_SOURCE/crates/litchi-ops/Cargo.toml" \
    --toolchain "$CANDIDATE_SOURCE/rust-toolchain.toml" \
    --source-assembly "$FINAL_ASSEMBLY" \
    --build-receipt "$W8_PREP/build-after.json" \
    --validation-receipt "$FINAL_VALIDATION_RECEIPT" \
    --output "$FINAL_MANIFEST" --source-commit "$FINAL_COMMIT" \
    --all-features --target "$TARGET" --profile "$PROFILE"
}

pair() {
  require_inputs
  verify_candidate_archive "$FINAL_OUT/472b-source.tar"
  python3 "$ROOT/tools/w8_projection_policy.py" pair \
    --policy "$POLICY" --before-manifest "$BASE_MANIFEST" \
    --after-manifest "$FINAL_MANIFEST" --output "$PAIR_RECEIPT"
}

abba() {
  require_inputs
  verify_candidate_archive "$FINAL_OUT/472b-source.tar"
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

if [[ "${BASH_SOURCE[0]}" == "$0" ]]; then
  case "${1:-}" in
    prepare|build|validate|pair|abba|allocator) "$1" ;;
    *) echo "usage: $0 {prepare|build|validate|pair|abba|allocator}" >&2; exit 2 ;;
  esac
fi
