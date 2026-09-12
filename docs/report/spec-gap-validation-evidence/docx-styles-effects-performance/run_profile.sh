#!/usr/bin/env bash
set -euo pipefail

# This runner is deliberately sealed until the scaffold has been reviewed and
# committed. The default is refusal; a caller must opt in explicitly after
# the reviewed commit, so a copied script cannot accidentally start a profile.
if [[ "${PROFILE_FROZEN:-}" != "1" ]]; then
    echo "profile runner is frozen; set PROFILE_FROZEN=1 only after review" >&2
    exit 2
fi
if [[ "${DOCX_STYLES_EFFECTS_PROFILE_API_WIRED:-}" != "1" ]]; then
    echo "profile API wiring is not authorized" >&2
    exit 2
fi
if [[ -n "${RUSTFLAGS:-}" || -n "${CARGO_ENCODED_RUSTFLAGS:-}" || -n "${RUSTC_BOOTSTRAP:-}" || -n "${RUSTDOCFLAGS:-}" ]]; then
    echo "profile refuses inherited Rust flags or bootstrap settings" >&2
    exit 2
fi
if [[ ! -x /usr/bin/time ]]; then
    echo "profile requires /usr/bin/time" >&2
    exit 2
fi

HERE=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd -P)
ROOT=$(cd -- "$HERE/../../../.." && pwd -P)
# Rustup and Cargo discover toolchain/config files from the working directory.
# Anchor their execution to this frozen checkout, regardless of the caller.
cd -- "$ROOT"
SOURCE_COMMIT=d000d977b99e03f8542c7dae74acf767a91b1feb
# The profile's correctness prerequisite is the separately reviewed current
# source smoke retained in c4ac5163. The historical d1/clean-46 smoke remains
# replayable through run_smoke.sh and verify.py, but must not gate this profile.
CURRENT_SMOKE_COMMIT=c4ac516353591f88a9b797349002d05737614e78
HARNESS=$HERE/profile-harness/Cargo.toml
MANIFEST_TOOL=$HERE/source_manifest.py
PROFILE_VERIFIER=$HERE/verify_profile.py
HOST_PROBE=$HERE/host_probe.sh

RESULTS=${DOCX_STYLES_EFFECTS_PROFILE_RESULTS:-}
TARGET=${DOCX_STYLES_EFFECTS_PROFILE_TARGET_DIR:-/var/tmp/litchi-docx-styles-effects-profile-target}
if [[ -z "$RESULTS" || "$RESULTS" != /* || "$RESULTS" == "/" ]]; then
    echo "DOCX_STYLES_EFFECTS_PROFILE_RESULTS must be a fresh absolute path" >&2
    exit 2
fi
if [[ "$TARGET" != /* || "$TARGET" == "/" ]]; then
    echo "DOCX_STYLES_EFFECTS_PROFILE_TARGET_DIR must be an absolute path" >&2
    exit 2
fi
if ! command -v realpath >/dev/null 2>&1; then
    echo "profile requires realpath for canonical path checks" >&2
    exit 2
fi
RESULTS=$(realpath -m -- "$RESULTS")
TARGET=$(realpath -m -- "$TARGET")
if [[ "$RESULTS" == "/" || "$TARGET" == "/" ]]; then
    echo "profile results and target cannot be the filesystem root" >&2
    exit 2
fi
case "$RESULTS/" in
    "$ROOT/"*) echo "profile results must be outside the checkout" >&2; exit 2 ;;
esac
case "$TARGET/" in
    "$ROOT/"*) echo "profile target must be outside the checkout" >&2; exit 2 ;;
esac
case "$RESULTS/" in
    "$TARGET/"*) echo "profile results and target must be disjoint" >&2; exit 2 ;;
esac
case "$TARGET/" in
    "$RESULTS/"*) echo "profile results and target must be disjoint" >&2; exit 2 ;;
esac
if [[ -e "$RESULTS" || -L "$RESULTS" || -e "$TARGET" || -L "$TARGET" ]]; then
    echo "profile results and target must be fresh paths" >&2
    exit 2
fi

PROCESSES=${DOCX_STYLES_EFFECTS_PROFILE_PROCESSES:-3}
WARMUP=${DOCX_STYLES_EFFECTS_PROFILE_WARMUP:-2}
SAMPLES=${DOCX_STYLES_EFFECTS_PROFILE_SAMPLES:-20}
[[ "$PROCESSES" =~ ^[0-9]+$ && "$PROCESSES" -eq 3 ]] || { echo "profile requires exactly 3 fresh processes" >&2; exit 2; }
[[ "$WARMUP" =~ ^[0-9]+$ && "$WARMUP" -eq 2 ]] || { echo "profile requires exactly 2 warmups" >&2; exit 2; }
[[ "$SAMPLES" =~ ^[0-9]+$ && "$SAMPLES" -eq 20 ]] || { echo "profile requires exactly 20 samples" >&2; exit 2; }

HEAD=$(git -C "$ROOT" rev-parse HEAD)
if ! git -C "$ROOT" merge-base --is-ancestor "$SOURCE_COMMIT" "$HEAD"; then
    echo "checkout does not descend from approved production source $SOURCE_COMMIT" >&2
    exit 2
fi
if ! git -C "$ROOT" merge-base --is-ancestor "$CURRENT_SMOKE_COMMIT" "$HEAD"; then
    echo "checkout does not descend from reviewed current-source correctness smoke $CURRENT_SMOKE_COMMIT" >&2
    exit 2
fi
if [[ -n "$(git -C "$ROOT" status --porcelain --untracked-files=all)" ]]; then
    echo "profile requires a clean committed checkout" >&2
    exit 2
fi

# Require the reviewed correctness capture before any profile build or
# staging. This prevents a profile from silently replacing its prerequisite.
python3 "$PROFILE_VERIFIER" --smoke-only --root "$ROOT" --evidence "$HERE"

REQUIRED_FILES=(
    "$HERE/profile-harness/Cargo.toml"
    "$HERE/profile-harness/Cargo.lock"
    "$HERE/profile-harness/main.rs"
    "$HERE/profile-harness/synthetic.rs"
    "$HERE/profile-harness/profile_adapter.rs"
    "$HERE/harness/Cargo.toml"
    "$HERE/harness/Cargo.lock"
    "$HERE/harness/main.rs"
    "$HERE/harness/adapter.rs"
    "$HERE/harness/support.rs"
    "$HERE/source_manifest.py"
    "$HERE/verify.py"
    "$HERE/verify_profile.py"
    "$HERE/test_source_snapshot.py"
    "$HERE/test_verify.py"
    "$HERE/test_profile_scaffold.py"
    "$HERE/corpus-manifest.json"
    "$HERE/requirements.md"
    "$HERE/performance-plan.md"
    "$HERE/README.md"
    "$HERE/run_smoke.sh"
    "$HERE/run_profile.sh"
    "$HERE/host_probe.sh"
    "$HERE/cleanup_target.sh"
    "$ROOT/crates/litchi-docx/src/styles/effects.rs"
    "$ROOT/crates/litchi-docx/src/package/package/styles_with_effects.rs"
    "$ROOT/crates/litchi-docx/tests/styles_with_effects.rs"
    "$ROOT/crates/litchi-opc/src/phys_pkg.rs"
    "$ROOT/crates/litchi-opc/src/limits.rs"
    "$HERE/fixtures/Bug54849.docx"
    "$HERE/fixtures/ComplexNumberedLists.docx"
    "$HERE/fixtures/ms-office-2010-signed.docx"
    "$HERE/fixtures/testGlossary.docx"
)
for required in "${REQUIRED_FILES[@]}"; do
    [[ -f "$required" ]] || { echo "required profile input is missing: $required" >&2; exit 2; }
    shown=${required#"$ROOT/"}
    git -C "$ROOT" ls-files --error-unmatch -- "$shown" >/dev/null || {
        echo "profile input is not committed: $shown" >&2
        exit 2
    }
done

mkdir -p -- "$(dirname -- "$RESULTS")"
mkdir -- "$RESULTS"
mkdir -p -- "$(dirname -- "$TARGET")"
mkdir -- "$TARGET"
# Install cleanup immediately after staging creation. Raw receipts remain in
# RESULTS for replay. A target is deleted only after the verifier writes a
# matching success sentinel; any failure retains it for diagnosis.
SUCCESS_SENTINEL=$RESULTS/profile-success.sentinel
source "$HERE/cleanup_target.sh"
trap cleanup_profile_target EXIT

export CARGO_TARGET_DIR="$TARGET"
export CARGO_INCREMENTAL=0
export LC_ALL=C
unset RUSTFLAGS CARGO_ENCODED_RUSTFLAGS RUSTC_BOOTSTRAP RUSTDOCFLAGS

COMMANDS=$RESULTS/commands.txt
BUILD_LOG=$RESULTS/build.log
METADATA_BEFORE=$RESULTS/metadata-before.json
METADATA_AFTER=$RESULTS/metadata-after.json
MANIFEST_BEFORE=$RESULTS/source-manifest-before.txt
MANIFEST_AFTER=$RESULTS/source-manifest-after.txt
HOST_BEFORE=$RESULTS/host-before.txt
HOST_AFTER=$RESULTS/host-after.txt
TOOLCHAIN_BEFORE=$RESULTS/toolchain-before.txt
GIT_STATUS_BEFORE=$RESULTS/git-status-before.txt
GIT_STATUS_AFTER=$RESULTS/git-status-after.txt
ENVIRONMENT_BEFORE=$RESULTS/environment-before.txt
BINARY=$TARGET/release/docx-styles-effects-profile
GENERATED_DIR=$RESULTS/generated-fixtures

: >"$COMMANDS"
printf 'source_commit=%s\n' "$SOURCE_COMMIT" >>"$COMMANDS"
printf 'git_head=%s\n' "$HEAD" >>"$COMMANDS"
printf 'environment=CARGO_TARGET_DIR=%s CARGO_INCREMENTAL=0 LC_ALL=C DOCX_PROFILE_GENERATED_FIXTURES=%s RUSTFLAGS=unset CARGO_ENCODED_RUSTFLAGS=unset RUSTC_BOOTSTRAP=unset RUSTDOCFLAGS=unset\n' "$TARGET" "$RESULTS/generated-fixtures.json" >>"$COMMANDS"

record_command() {
    local label=$1
    shift
    printf '%s=' "$label" >>"$COMMANDS"
    printf '%q ' "$@" >>"$COMMANDS"
    printf '\n' >>"$COMMANDS"
}

record_command command-verify-smoke python3 "$PROFILE_VERIFIER" --smoke-only --root "$ROOT" --evidence "$HERE"
python3 "$PROFILE_VERIFIER" --smoke-only --root "$ROOT" --evidence "$HERE"

record_command command-host-before bash "$HOST_PROBE" "$HOST_BEFORE"
bash "$HOST_PROBE" "$HOST_BEFORE"

git -C "$ROOT" status --porcelain --untracked-files=all >"$GIT_STATUS_BEFORE"
env | LC_ALL=C sort | grep -E '^(CARGO_TARGET_DIR|CARGO_INCREMENTAL|LC_ALL|RUSTFLAGS|CARGO_ENCODED_RUSTFLAGS|RUSTC_BOOTSTRAP|RUSTDOCFLAGS)=' >"$ENVIRONMENT_BEFORE" || true

record_command command-toolchain-before bash -c 'rustc --version; rustc -vV; cargo --version; rustup show active-toolchain'
TARGET_TRIPLE=$(rustc -vV | awk '/^host:/ {print $2}')
LINKER_ENV=CARGO_TARGET_${TARGET_TRIPLE^^}
LINKER_ENV=${LINKER_ENV//-/_}_LINKER
LINKER_VALUE=${!LINKER_ENV:-default}
{
    rustc --version
    printf 'rustc -vV\n'
    rustc -vV
    printf 'cargo --version\n'
    cargo --version
    printf 'active-toolchain: '
    rustup show active-toolchain
    printf 'target_triple=%s\n' "$TARGET_TRIPLE"
    printf 'linker=%s\n' "$LINKER_VALUE"
} >"$TOOLCHAIN_BEFORE"

record_command command-cargo-metadata-before cargo metadata --locked --offline --manifest-path "$HARNESS" --format-version 1
cargo metadata --locked --offline --manifest-path "$HARNESS" --format-version 1 >"$METADATA_BEFORE"

EXTRAS=(
    "$ROOT/Cargo.toml"
    "$ROOT/rust-toolchain.toml"
    "$ROOT/.cargo/config.toml"
    "$HERE/requirements.md"
    "$HERE/performance-plan.md"
    "$HERE/corpus-manifest.json"
    "$HERE/README.md"
    "$HERE/run_smoke.sh"
    "$HERE/run_profile.sh"
    "$HERE/host_probe.sh"
    "$HERE/cleanup_target.sh"
    "$HERE/source_manifest.py"
    "$HERE/verify.py"
    "$HERE/verify_profile.py"
    "$HERE/test_source_snapshot.py"
    "$HERE/test_verify.py"
    "$HERE/test_profile_scaffold.py"
    "$HERE/profile-harness/Cargo.toml"
    "$HERE/profile-harness/Cargo.lock"
    "$HERE/profile-harness/main.rs"
    "$HERE/profile-harness/synthetic.rs"
    "$HERE/profile-harness/profile_adapter.rs"
    "$HERE/harness/Cargo.toml"
    "$HERE/harness/Cargo.lock"
    "$HERE/harness/main.rs"
    "$HERE/harness/adapter.rs"
    "$HERE/harness/support.rs"
    "$ROOT/crates/litchi-docx/src/styles/effects.rs"
    "$ROOT/crates/litchi-docx/src/package/package/styles_with_effects.rs"
    "$ROOT/crates/litchi-docx/tests/styles_with_effects.rs"
    "$ROOT/crates/litchi-opc/src/phys_pkg.rs"
    "$ROOT/crates/litchi-opc/src/limits.rs"
    "$HERE/fixtures/Bug54849.docx"
    "$HERE/fixtures/ComplexNumberedLists.docx"
    "$HERE/fixtures/ms-office-2010-signed.docx"
    "$HERE/fixtures/testGlossary.docx"
)
MANIFEST_ARGS=(python3 "$MANIFEST_TOOL" --metadata "$METADATA_BEFORE" --root "$ROOT" --evidence "$HERE" --output "$MANIFEST_BEFORE" --source-commit "$SOURCE_COMMIT")
for extra in "${EXTRAS[@]}"; do
    MANIFEST_ARGS+=(--extra "$extra")
done
record_command command-source-manifest-before "${MANIFEST_ARGS[@]}"
"${MANIFEST_ARGS[@]}"

record_command command-cargo-build cargo build --release --locked --offline --manifest-path "$HARNESS"
cargo build --release --locked --offline --manifest-path "$HARNESS" >"$BUILD_LOG" 2>&1
[[ -x "$BINARY" ]] || { echo "profile binary was not produced: $BINARY" >&2; exit 1; }
sha256sum "$BINARY" >"$RESULTS/binary.sha256"

RUSTC_VERSION=$(rustc --version)
CARGO_VERSION=$(cargo --version)
{
    sha256sum "$BINARY"
    printf 'binary=%s\n' "$BINARY"
    printf 'source_commit=%s\n' "$SOURCE_COMMIT"
    printf 'git_head=%s\n' "$HEAD"
    printf 'git_status_before_sha256=%s\n' "$(sha256sum "$GIT_STATUS_BEFORE" | awk '{print $1}')"
    printf 'git_status_after_sha256=unavailable\n'
    printf 'rustc=%s\n' "$RUSTC_VERSION"
    printf 'cargo=%s\n' "$CARGO_VERSION"
    printf 'target=%s\n' "$TARGET"
    printf 'target_triple=%s\n' "$TARGET_TRIPLE"
    printf 'linker=%s\n' "$LINKER_VALUE"
    printf 'toolchain_sha256=%s\n' "$(sha256sum "$TOOLCHAIN_BEFORE" | awk '{print $1}')"
    printf 'environment_sha256=%s\n' "$(sha256sum "$COMMANDS" | awk '{print $1}')"
    printf 'allocator=profile-harness::support::CountingAllocator\n'
    printf 'rss=/usr/bin/time -v process maximum resident set size\n'
    printf 'mode=profile processes=%s warmup=%s samples=%s\n' "$PROCESSES" "$WARMUP" "$SAMPLES"
} >"$RESULTS/build-provenance.txt"

mkdir -- "$GENERATED_DIR"
record_command command-fixture-manifest "$BINARY" --emit-fixture-manifest --output-dir "$GENERATED_DIR"
"$BINARY" --emit-fixture-manifest --output-dir "$GENERATED_DIR" >"$RESULTS/generated-fixtures.json"
export DOCX_PROFILE_GENERATED_FIXTURES="$RESULTS/generated-fixtures.json"
env | LC_ALL=C sort | grep -E '^(CARGO_TARGET_DIR|CARGO_INCREMENTAL|LC_ALL|DOCX_PROFILE_GENERATED_FIXTURES|RUSTFLAGS|CARGO_ENCODED_RUSTFLAGS|RUSTC_BOOTSTRAP|RUSTDOCFLAGS)=' >"$ENVIRONMENT_BEFORE" || true

specs=(
    'capture_bug_main|native'
    'capture_bug_glossary|native'
    'capture_signed_main|native'
    'capture_complex_main|native'
    'capture_glossary_main|native'
    'capture_glossary_glossary|native'
    'capture_synthetic_main|64k'
    'capture_synthetic_main|1m'
    'capture_synthetic_glossary|64k'
    'capture_synthetic_glossary|1m'
    'projection_main|native'
    'projection_main|64k'
    'projection_main|1m'
    'projection_glossary|native'
    'projection_glossary|64k'
    'projection_glossary|1m'
    'noop_main|native'
    'noop_glossary|native'
    'replace_main|native'
    'replace_main|64k'
    'replace_main|1m'
    'replace_glossary|native'
    'replace_glossary|64k'
    'replace_glossary|1m'
    'remove_main|native'
    'remove_glossary|native'
    'add_main_absent|native'
    'inverse_replace_main|native'
    'inverse_remove_main|native'
    'independent_main|native'
    'independent_glossary|native'
)

for ((process=1; process<=PROCESSES; process++)); do
    for spec in "${specs[@]}"; do
        lane=${spec%%|*}
        scale=${spec#*|}
        stem=profile-${lane}-${scale}-p${process}
        receipt=$RESULTS/$stem.json
        stderr=$RESULTS/$stem.stderr.log
        timing=$RESULTS/$stem.time.txt
        printf -v run_command '%q ' /usr/bin/time -v -o "$timing" "$BINARY" --lane "$lane" --scale "$scale" --warmup "$WARMUP" --samples "$SAMPLES"
        printf 'run=%sfresh_process=1 process_index=%s lane=%s scale=%s\n' "$run_command" "$process" "$lane" "$scale" >>"$COMMANDS"
        DOCX_PROFILE_PROCESS_INDEX=$process /usr/bin/time -v -o "$timing" "$BINARY" --lane "$lane" --scale "$scale" --warmup "$WARMUP" --samples "$SAMPLES" >"$receipt" 2>"$stderr"
    done
done

git -C "$ROOT" status --porcelain --untracked-files=all >"$GIT_STATUS_AFTER"
if ! cmp -s "$GIT_STATUS_BEFORE" "$GIT_STATUS_AFTER"; then
    echo "profile checkout changed during capture; target retained for diagnosis" >&2
    exit 1
fi

record_command command-cargo-metadata-after cargo metadata --locked --offline --manifest-path "$HARNESS" --format-version 1
cargo metadata --locked --offline --manifest-path "$HARNESS" --format-version 1 >"$METADATA_AFTER"
MANIFEST_ARGS_AFTER=(python3 "$MANIFEST_TOOL" --metadata "$METADATA_AFTER" --root "$ROOT" --evidence "$HERE" --output "$MANIFEST_AFTER" --source-commit "$SOURCE_COMMIT")
for extra in "${EXTRAS[@]}"; do
    MANIFEST_ARGS_AFTER+=(--extra "$extra")
done
record_command command-source-manifest-after "${MANIFEST_ARGS_AFTER[@]}"
"${MANIFEST_ARGS_AFTER[@]}"
sha256sum "$BINARY" >"$RESULTS/binary-after.sha256"

{
    sha256sum "$BINARY"
    printf 'binary=%s\n' "$BINARY"
    printf 'source_commit=%s\n' "$SOURCE_COMMIT"
    printf 'git_head=%s\n' "$HEAD"
    printf 'git_status_before_sha256=%s\n' "$(sha256sum "$GIT_STATUS_BEFORE" | awk '{print $1}')"
    printf 'git_status_after_sha256=%s\n' "$(sha256sum "$GIT_STATUS_AFTER" | awk '{print $1}')"
    printf 'rustc=%s\n' "$RUSTC_VERSION"
    printf 'cargo=%s\n' "$CARGO_VERSION"
    printf 'target=%s\n' "$TARGET"
    printf 'target_triple=%s\n' "$TARGET_TRIPLE"
    printf 'linker=%s\n' "$LINKER_VALUE"
    printf 'toolchain_sha256=%s\n' "$(sha256sum "$TOOLCHAIN_BEFORE" | awk '{print $1}')"
    printf 'environment_sha256=%s\n' "$(sha256sum "$ENVIRONMENT_BEFORE" | awk '{print $1}')"
    printf 'allocator=profile-harness::support::CountingAllocator\n'
    printf 'rss=/usr/bin/time -v process maximum resident set size\n'
    printf 'mode=profile processes=%s warmup=%s samples=%s\n' "$PROCESSES" "$WARMUP" "$SAMPLES"
} >"$RESULTS/build-provenance.txt"

{
    printf 'source_manifest_sha256=%s\n' "$(sha256sum "$MANIFEST_BEFORE" | awk '{print $1}')"
    printf 'metadata_before_sha256=%s\n' "$(sha256sum "$METADATA_BEFORE" | awk '{print $1}')"
    printf 'metadata_after_sha256=%s\n' "$(sha256sum "$METADATA_AFTER" | awk '{print $1}')"
    printf 'git_head=%s\n' "$HEAD"
    printf 'git_status_before_sha256=%s\n' "$(sha256sum "$GIT_STATUS_BEFORE" | awk '{print $1}')"
    printf 'git_status_after_sha256=%s\n' "$(sha256sum "$GIT_STATUS_AFTER" | awk '{print $1}')"
} >"$RESULTS/source-provenance.txt"

record_command command-host-after bash "$HOST_PROBE" "$HOST_AFTER"
bash "$HOST_PROBE" "$HOST_AFTER"

record_command command-verify-profile python3 "$PROFILE_VERIFIER" --root "$ROOT" --evidence "$HERE" --results "$RESULTS" --manifest-before "$MANIFEST_BEFORE" --manifest-after "$MANIFEST_AFTER" --metadata-before "$METADATA_BEFORE" --metadata-after "$METADATA_AFTER" --output "$RESULTS/verification.json"
python3 "$PROFILE_VERIFIER" --root "$ROOT" --evidence "$HERE" --results "$RESULTS" --manifest-before "$MANIFEST_BEFORE" --manifest-after "$MANIFEST_AFTER" --metadata-before "$METADATA_BEFORE" --metadata-after "$METADATA_AFTER" --output "$RESULTS/verification.json"
printf 'verification_sha256=%s\ngit_head=%s\ntarget=%s\n' "$(sha256sum "$RESULTS/verification.json" | awk '{print $1}')" "$HEAD" "$TARGET" >"$SUCCESS_SENTINEL"

printf 'profile verification passed: %s\n' "$RESULTS/verification.json"
