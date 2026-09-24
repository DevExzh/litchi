#!/usr/bin/env bash

# Run the frozen litchi-xlsb library gates and retain reproducible evidence.
# The caller must set FINAL_GATE_FREEZE=1 after the reviewer/coder source
# freeze. This guard prevents an accidental broad build during active edits.

set -euo pipefail

repo_root=$(git rev-parse --show-toplevel)
gates_dir=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
target_dir=${CARGO_TARGET_DIR:-/var/tmp/litchi-xlsb-drawing-current-target-20260910}
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

source_paths=(
    crates/litchi-xlsb/src/binary_index.rs
    crates/litchi-xlsb/src/worksheet_index.rs
    crates/litchi-xlsb/src/lib.rs
    crates/litchi-xlsb/src/cell_values/drawing_transfer.rs
    crates/litchi-xlsb/src/cell_values/root.rs
    crates/litchi-xlsb/src/cell_values/workbook.rs
    crates/litchi-xlsb/src/cell_watches/workbook.rs
    crates/litchi-xlsb/src/raw/kind.rs
    crates/litchi-xlsb/src/slicer/package.rs
    crates/litchi-xlsb/src/sparkline/workbook.rs
    crates/litchi-xlsb/src/timeline/package.rs
    crates/litchi-xlsb/src/workbook/codec/package.rs
    crates/litchi-xlsb/src/workbook/source.rs
    crates/litchi-xlsb/src/workbook/package.rs
    crates/litchi-xlsb/src/workbook/mod.rs
    crates/litchi-xlsb/src/cell_values/worksheet.rs
    crates/litchi-xlsb/src/writer/workbook/package.rs
    crates/litchi-xlsb/src/writer/workbook/model.rs
    crates/litchi-xlsb/src/xml_maps/patch.rs
    crates/litchi-xlsb/tests/worksheet_binary_index.rs
    crates/litchi-xlsb/examples/binary_index_profile.rs
)

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
tests: cargo +$toolchain test -p litchi-xlsb --all-features --all-targets --offline --locked --no-fail-fast
doctests: cargo +$toolchain test -p litchi-xlsb --all-features --doc --offline --locked --no-fail-fast
clippy: cargo +$toolchain clippy -p litchi-xlsb --all-features --all-targets --offline --locked -- -D warnings
rustdoc: RUSTDOCFLAGS=-Dwarnings cargo +$toolchain doc -p litchi-xlsb --all-features --no-deps --offline --locked
format: cargo +$toolchain fmt -p litchi-xlsb -- --check
diff: git diff HEAD --check
EOF

cargo "+$toolchain" test -p litchi-xlsb --all-features --all-targets --offline --locked --no-fail-fast \
    >"$gates_dir/xlsb-all-features-all-targets-incremental0.log" 2>&1
cargo "+$toolchain" test -p litchi-xlsb --all-features --doc --offline --locked --no-fail-fast \
    >"$gates_dir/xlsb-all-features-doctests-incremental0.log" 2>&1
cargo "+$toolchain" clippy -p litchi-xlsb --all-features --all-targets --offline --locked -- -D warnings \
    >"$gates_dir/xlsb-clippy-all-features-all-targets-incremental0.log" 2>&1
RUSTDOCFLAGS=-Dwarnings cargo "+$toolchain" doc -p litchi-xlsb --all-features --no-deps --offline --locked \
    >"$gates_dir/xlsb-rustdoc-all-features.log" 2>&1
cargo "+$toolchain" fmt -p litchi-xlsb -- --check \
    >"$gates_dir/xlsb-format.log" 2>&1
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
