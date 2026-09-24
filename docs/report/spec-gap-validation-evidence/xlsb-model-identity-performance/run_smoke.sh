#!/usr/bin/env bash
set -euo pipefail

if [[ -n "${RUSTFLAGS:-}" || -n "${CARGO_ENCODED_RUSTFLAGS:-}" || -n "${RUSTC_BOOTSTRAP:-}" ]]; then
    echo "smoke runner refuses nonempty compiler override flags" >&2
    exit 2
fi

script_dir="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
root_dir="$(cd -- "$script_dir/../../../.." && pwd)"
profile_dir="$script_dir"
harness_manifest="$profile_dir/harness/Cargo.toml"
manifest_tool="$profile_dir/source_manifest.py"
result_dir="$(realpath -m -- "${XLSB_PROFILE_RESULT_DIR:-$profile_dir/results/smoke-fixed}")"
current_commit="$(git -C "$root_dir" --no-replace-objects rev-parse HEAD)"
host_commit="1b0d4864804d666aaf4ae5039bd6300996ca32a5"
if ! git -C "$root_dir" --no-replace-objects merge-base --is-ancestor "$host_commit" "$current_commit"; then
    echo "refusing smoke without the pinned host implementation commit" >&2
    exit 2
fi
target_input="${XLSB_PROFILE_TARGET_DIR:-}"
if [[ -n "$target_input" ]]; then
    target_dir="$(realpath -m -- "$target_input")"
    if [[ -e "$target_dir" ]]; then
        echo "refusing to reuse an existing target: $target_dir" >&2
        exit 2
    fi
    mkdir -p "$target_dir"
else
    target_dir="$(mktemp -d /var/tmp/litchi-xlsb-model-identity-profile.XXXXXX)"
fi
target_created=1

cleanup() {
    status=$?
    if [[ "$target_created" == 1 && -d "$target_dir" ]]; then
        find "$target_dir" -depth -delete
    fi
    exit "$status"
}
trap cleanup EXIT

mkdir -p "$result_dir"
if [[ -n "$(find "$result_dir" -mindepth 1 -maxdepth 1 -print -quit)" ]]; then
    echo "smoke result directory is not fresh: $result_dir" >&2
    echo "set XLSB_PROFILE_RESULT_DIR to a new directory for another run" >&2
    exit 2
fi

{
    echo "git_head=$current_commit"
    echo "git_baseline=$(git -C "$root_dir" --no-replace-objects rev-parse "$host_commit^{commit}")"
    echo "neutral_baseline=$(git -C "$root_dir" --no-replace-objects rev-parse 4f53be2d3^{commit})"
    echo "rustc=$(rustc -Vv)"
    echo "cargo=$(cargo -V)"
    echo "target_dir=$target_dir"
    echo "cargo_incremental=0"
    echo "uname=$(uname -a)"
} >"$result_dir/provenance.txt"

export CARGO_INCREMENTAL=0
export LC_ALL=C
unset RUSTFLAGS CARGO_ENCODED_RUSTFLAGS RUSTC_BOOTSTRAP RUSTDOCFLAGS

cargo metadata --format-version=1 --locked --offline --manifest-path "$harness_manifest" \
    >"$result_dir/metadata-before.json"
manifest_args=(
    --metadata "$result_dir/metadata-before.json"
    --root "$root_dir"
    --output "$result_dir/source-manifest-before.txt"
    --git-commit "$current_commit"
)
for extra in \
    "$root_dir/Cargo.toml" \
    "$root_dir/rust-toolchain.toml" \
    "$root_dir/.cargo/config.toml" \
    "$root_dir/crates/litchi-xlsb/tests/data_model_identity.rs" \
    "$profile_dir/requirements.md" \
    "$profile_dir/corpus-manifest.json" \
    "$profile_dir/README.md" \
    "$profile_dir/run_smoke.sh" \
    "$profile_dir/verify.py" \
    "$profile_dir/source_manifest.py" \
    "$profile_dir/test_provenance.py" \
    "$harness_manifest" \
    "$profile_dir/harness/Cargo.lock" \
    "$profile_dir/harness/build.rs" \
    "$profile_dir/harness/matrix.rs" \
    "$profile_dir/harness/scaled_fixture.rs" \
    "$profile_dir/harness/adapter.rs" \
    "$profile_dir/harness/main.rs" \
    "$profile_dir/harness/support.rs"; do
    manifest_args+=(--extra "$extra")
done
python3 "$manifest_tool" "${manifest_args[@]}"

CARGO_TARGET_DIR="$target_dir" cargo build \
    --manifest-path "$harness_manifest" \
    --offline --locked >"$result_dir/build.log" 2>&1

binary="$target_dir/debug/xlsb-model-identity-profile"
sha256sum "$binary" >"$result_dir/binary.sha256"
{
    echo "binary=$binary"
    cat "$result_dir/binary.sha256"
    echo "rustc -vV:"
    rustc -vV
    echo "cargo=$(cargo -V)"
    echo "target=$target_dir"
    echo "cargo_incremental=$CARGO_INCREMENTAL"
    echo "flags=none"
    echo "allocator=CountingAllocator (process-local GlobalAlloc observer)"
} >"$result_dir/build-provenance.txt"
lanes=(
    neutral_open_tiny
    host_open_tiny
    neutral_open_relationship
    host_stage_noop_tiny
    host_stage_rename_relationship
    host_commit_rename_relationship
    host_save_reopen_relationship
    host_inverse_relationship
    host_exact_cap_relationship
    host_refusal_opaque
    host_refusal_limit
)

for lane in "${lanes[@]}"; do
    /usr/bin/time -v -o "$result_dir/$lane.time.txt" \
        "$binary" --lane "$lane" --warmup 0 --samples 1 \
        >"$result_dir/$lane.json" \
        2>"$result_dir/$lane.stderr.log"
done

# The correctness matrix uses this exact built binary but has its own receipt
# and verifier. It intentionally runs outside /usr/bin/time: the 120-point
# matrix records semantic/byte gates only and makes no timing or allocator
# claim. Keep the argv, stdout, stderr, and exit status as independent
# artifacts so a later clean review can reproduce the exact invocation.
XLSB_PROFILE_MATRIX_BINARY="$binary" python3 - <<'PY' \
    >"$result_dir/matrix-correctness.argv.json"
import json
import os

print(json.dumps({"argv": [os.environ["XLSB_PROFILE_MATRIX_BINARY"], "--matrix-correctness"]}))
PY
set +e
"$binary" --matrix-correctness \
    >"$result_dir/matrix-correctness.stdout.json" \
    2>"$result_dir/matrix-correctness.stderr.log"
matrix_status=$?
set -e
printf '%s\n' "$matrix_status" >"$result_dir/matrix-correctness.exit.txt"
if [[ "$matrix_status" -ne 0 ]]; then
    echo "correctness matrix failed with exit status $matrix_status" >&2
    exit "$matrix_status"
fi

sha256sum "$binary" >"$result_dir/binary-after.sha256"
cmp -s "$result_dir/binary.sha256" "$result_dir/binary-after.sha256"

cargo metadata --format-version=1 --locked --offline --manifest-path "$harness_manifest" \
    >"$result_dir/metadata-after.json"
manifest_args_after=(
    --metadata "$result_dir/metadata-after.json"
    --root "$root_dir"
    --output "$result_dir/source-manifest-after.txt"
    --git-commit "$current_commit"
)
for extra in \
    "$root_dir/Cargo.toml" \
    "$root_dir/rust-toolchain.toml" \
    "$root_dir/.cargo/config.toml" \
    "$root_dir/crates/litchi-xlsb/tests/data_model_identity.rs" \
    "$profile_dir/requirements.md" \
    "$profile_dir/corpus-manifest.json" \
    "$profile_dir/README.md" \
    "$profile_dir/run_smoke.sh" \
    "$profile_dir/verify.py" \
    "$profile_dir/source_manifest.py" \
    "$profile_dir/test_provenance.py" \
    "$harness_manifest" \
    "$profile_dir/harness/Cargo.lock" \
    "$profile_dir/harness/build.rs" \
    "$profile_dir/harness/matrix.rs" \
    "$profile_dir/harness/scaled_fixture.rs" \
    "$profile_dir/harness/adapter.rs" \
    "$profile_dir/harness/main.rs" \
    "$profile_dir/harness/support.rs"; do
    manifest_args_after+=(--extra "$extra")
done
python3 "$manifest_tool" "${manifest_args_after[@]}"
cmp -s "$result_dir/source-manifest-before.txt" "$result_dir/source-manifest-after.txt"

{
    echo "metadata_before_sha256=$(sha256sum "$result_dir/metadata-before.json" | cut -d' ' -f1)"
    echo "metadata_after_sha256=$(sha256sum "$result_dir/metadata-after.json" | cut -d' ' -f1)"
    echo "source_manifest_before_sha256=$(sha256sum "$result_dir/source-manifest-before.txt" | cut -d' ' -f1)"
    echo "source_manifest_after_sha256=$(sha256sum "$result_dir/source-manifest-after.txt" | cut -d' ' -f1)"
    echo "binary_sha256=$(cut -d' ' -f1 "$result_dir/binary.sha256")"
    echo "binary_after_sha256=$(cut -d' ' -f1 "$result_dir/binary-after.sha256")"
    echo "git_status_relevant=$(git -C "$root_dir" status --short --untracked-files=all -- crates | tr '\n' ';')"
} >>"$result_dir/provenance.txt"

python3 "$profile_dir/verify.py" \
    --results "$result_dir" \
    --manifest "$profile_dir/corpus-manifest.json" \
    --root "$root_dir" \
    --source-before "$result_dir/source-manifest-before.txt" \
    --source-after "$result_dir/source-manifest-after.txt"

python3 "$profile_dir/verify.py" \
    --correctness-receipt "$result_dir/matrix-correctness.stdout.json" \
    --manifest "$profile_dir/corpus-manifest.json" \
    --root "$root_dir"
