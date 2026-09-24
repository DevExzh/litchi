#!/usr/bin/env bash

# Run the frozen XLSB and shared DrawingML theme gates and retain reproducible evidence.
# The caller must set FINAL_GATE_FREEZE=1 after the reviewer/coder source
# freeze. This guard prevents an accidental broad build during active edits.

set -euo pipefail

repo_root=$(git rev-parse --show-toplevel)
gates_dir=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
target_dir=${CARGO_TARGET_DIR:-/tmp/litchi-xlsb-theme-target-20260910}
toolchain=${RUST_TOOLCHAIN:-1.95.0}
build_jobs=${CARGO_BUILD_JOBS:-2}

if [[ "${FINAL_GATE_FREEZE:-}" != 1 ]]; then
    echo "refusing final gates: set FINAL_GATE_FREEZE=1 after source freeze" >&2
    exit 2
fi
if [[ ! "$build_jobs" =~ ^[1-9][0-9]*$ ]]; then
    echo "CARGO_BUILD_JOBS must be a positive decimal integer" >&2
    exit 2
fi

cd "$repo_root"
mkdir -p "$gates_dir"

export RUSTUP_HOME=${RUSTUP_HOME:-/tmp/litchi-spec-gap-rustup}
export CARGO_TARGET_DIR="$target_dir"
export CARGO_PROFILE_DEV_DEBUG=0
export CARGO_PROFILE_TEST_DEBUG=0
export CARGO_INCREMENTAL=0
export CARGO_BUILD_JOBS="$build_jobs"

mapfile -t source_paths <"$gates_dir/source-paths.txt"

for path in "${source_paths[@]}"; do
    [[ -f "$path" ]] || {
        echo "source manifest path does not exist: $path" >&2
        exit 1
    }
done

hash_sources() {
    sha256sum "${source_paths[@]}"
}

hash_sources >"$gates_dir/source-hashes-before.txt"
git rev-parse HEAD >"$gates_dir/git-rev-before.txt"
cp Cargo.lock "$gates_dir/workspace-Cargo.lock"
sha256sum Cargo.lock >"$gates_dir/workspace-Cargo.lock.sha256"

cat >"$gates_dir/commands.txt" <<EOF
toolchain: cargo +$toolchain
target: $CARGO_TARGET_DIR
environment: CARGO_PROFILE_DEV_DEBUG=$CARGO_PROFILE_DEV_DEBUG CARGO_PROFILE_TEST_DEBUG=$CARGO_PROFILE_TEST_DEBUG CARGO_INCREMENTAL=$CARGO_INCREMENTAL CARGO_BUILD_JOBS=$CARGO_BUILD_JOBS RUSTUP_HOME=$RUSTUP_HOME
tests: cargo +$toolchain test -p litchi-xlsb -p litchi-drawingml --all-features --all-targets --offline --locked --no-fail-fast
doctests: cargo +$toolchain test -p litchi-xlsb -p litchi-drawingml --all-features --doc --offline --locked --no-fail-fast
pptx-consumer-unit: cargo +$toolchain test -p litchi-pptx --lib shape::theme --offline --locked
pptx-consumer-integration: cargo +$toolchain test -p litchi-pptx --test pptx_master_themes --offline --locked
schema-fixtures: LITCHI_THEME_SCHEMA_OUTPUT_DIR=$gates_dir/../generated-xml cargo +$toolchain test -p litchi-xlsb --all-features --lib theme::splice_tests --offline --locked
schema-validation: python3 $gates_dir/../verify-theme-schema.py $gates_dir/../generated-xml/*.xml
clippy: cargo +$toolchain clippy -p litchi-xlsb -p litchi-drawingml --all-features --all-targets --offline --locked -- -D warnings
rustdoc: RUSTDOCFLAGS=-Dwarnings cargo +$toolchain doc -p litchi-xlsb -p litchi-drawingml --all-features --no-deps --offline --locked
format: cargo +$toolchain fmt -p litchi-xlsb -p litchi-drawingml -- --check
diff: git diff HEAD --check
EOF

cargo "+$toolchain" test -p litchi-xlsb -p litchi-drawingml --all-features --all-targets --offline --locked --no-fail-fast \
    >"$gates_dir/theme-all-features-all-targets-incremental0.log" 2>&1
cargo "+$toolchain" test -p litchi-xlsb -p litchi-drawingml --all-features --doc --offline --locked --no-fail-fast \
    >"$gates_dir/theme-all-features-doctests-incremental0.log" 2>&1
cargo "+$toolchain" test -p litchi-pptx --lib shape::theme --offline --locked \
    >"$gates_dir/theme-pptx-consumer-unit.log" 2>&1
cargo "+$toolchain" test -p litchi-pptx --test pptx_master_themes --offline --locked \
    >"$gates_dir/theme-pptx-consumer-integration.log" 2>&1
LITCHI_THEME_SCHEMA_OUTPUT_DIR="$gates_dir/../generated-xml" \
    cargo "+$toolchain" test -p litchi-xlsb --all-features --lib theme::splice_tests --offline --locked \
    >"$gates_dir/theme-generated-schema-fixtures.log" 2>&1
python3 "$gates_dir/../verify-theme-schema.py" "$gates_dir"/../generated-xml/*.xml \
    >"$gates_dir/../generated-schema-validation.json"
cargo "+$toolchain" clippy -p litchi-xlsb -p litchi-drawingml --all-features --all-targets --offline --locked -- -D warnings \
    >"$gates_dir/theme-clippy-all-features-all-targets-incremental0.log" 2>&1
RUSTDOCFLAGS=-Dwarnings cargo "+$toolchain" doc -p litchi-xlsb -p litchi-drawingml --all-features --no-deps --offline --locked \
    >"$gates_dir/theme-rustdoc-all-features.log" 2>&1
cargo "+$toolchain" fmt -p litchi-xlsb -p litchi-drawingml -- --check \
    >"$gates_dir/theme-format.log" 2>&1
git diff HEAD --check >"$gates_dir/diff.log" 2>&1

hash_sources >"$gates_dir/source-hashes-after.txt"
git rev-parse HEAD >"$gates_dir/git-rev-after.txt"
if ! cmp -s "$gates_dir/source-hashes-before.txt" "$gates_dir/source-hashes-after.txt"; then
    echo "source changed while final gates were running" >&2
    exit 1
fi
if ! cmp -s "$gates_dir/git-rev-before.txt" "$gates_dir/git-rev-after.txt"; then
    echo "git HEAD changed while final gates were running" >&2
    exit 1
fi

printf 'toolchain=%s\n' "$toolchain" >"$gates_dir/gate-environment.txt"
printf 'rustc=' >>"$gates_dir/gate-environment.txt"
rustc +"$toolchain" --version >>"$gates_dir/gate-environment.txt"
printf 'cargo=' >>"$gates_dir/gate-environment.txt"
cargo +"$toolchain" --version >>"$gates_dir/gate-environment.txt"
printf 'target_dir=%s\n' "$CARGO_TARGET_DIR" >>"$gates_dir/gate-environment.txt"
printf 'CARGO_PROFILE_DEV_DEBUG=%s\n' "$CARGO_PROFILE_DEV_DEBUG" >>"$gates_dir/gate-environment.txt"
printf 'CARGO_PROFILE_TEST_DEBUG=%s\n' "$CARGO_PROFILE_TEST_DEBUG" >>"$gates_dir/gate-environment.txt"
printf 'CARGO_INCREMENTAL=%s\n' "$CARGO_INCREMENTAL" >>"$gates_dir/gate-environment.txt"
printf 'CARGO_BUILD_JOBS=%s\n' "$CARGO_BUILD_JOBS" >>"$gates_dir/gate-environment.txt"
printf 'RUSTUP_HOME=%s\n' "$RUSTUP_HOME" >>"$gates_dir/gate-environment.txt"

echo "library gates passed; evidence written to $gates_dir"
