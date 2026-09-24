#!/usr/bin/env bash

# Run one XLSB Theme case per fresh process.  The source and eager lanes are
# retained as separate API scopes; their timings are descriptive and do not
# imply an equivalent-work speedup.

set -euo pipefail

repo_root=$(git rev-parse --show-toplevel)
script_dir=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
output_dir=${OUTPUT_DIR:-"$script_dir/raw"}
target_dir=${CARGO_TARGET_DIR:-"$repo_root/target"}
warmup=${WARMUP:-3}
samples=${SAMPLES:-30}
processes=${PROCESSES:-1}
toolchain=${RUST_TOOLCHAIN:-1.95.0}
build_jobs=${CARGO_BUILD_JOBS:-2}
control_bytes=${THEME_CONTROL_BYTES:-1048576}
export RUSTUP_HOME=${RUSTUP_HOME:-/tmp/litchi-spec-gap-rustup}
export CARGO_INCREMENTAL=0

if [[ "$target_dir" != /* ]]; then
    target_dir="$repo_root/$target_dir"
fi
if [[ -n "${PROFILE_BIN:-}" ]]; then
    profile_bin=$PROFILE_BIN
    if [[ "$profile_bin" != /* ]]; then
        profile_bin="$repo_root/$profile_bin"
    fi
else
    profile_bin="$target_dir/release/examples/theme_profile"
fi

[[ "$processes" =~ ^[1-9][0-9]*$ ]] || { echo "PROCESSES must be positive" >&2; exit 1; }
[[ "$build_jobs" =~ ^[1-9][0-9]*$ ]] || { echo "CARGO_BUILD_JOBS must be positive" >&2; exit 1; }
(( build_jobs <= 2 )) || { echo "CARGO_BUILD_JOBS must not exceed 2" >&2; exit 1; }
[[ "$control_bytes" =~ ^[1-9][0-9]*$ ]] || { echo "THEME_CONTROL_BYTES must be positive" >&2; exit 1; }
(( control_bytes <= 7 * 1024 * 1024 )) || {
    echo "THEME_CONTROL_BYTES must stay below the 8 MiB Theme XML ceiling" >&2
    exit 1
}
mkdir -p "$output_dir"

source_paths=(
    crates/litchi-drawingml/src/theme/codec.rs
    crates/litchi-drawingml/src/theme/model.rs
    crates/litchi-xlsb/src/lib.rs
    crates/litchi-xlsb/src/theme.rs
    crates/litchi-xlsb/src/theme/lifecycle.rs
    crates/litchi-xlsb/src/workbook/source.rs
    crates/litchi-xlsb/src/workbook/package.rs
    crates/litchi-xlsb/examples/theme_profile.rs
    crates/litchi-xlsb/tests/theme.rs
    crates/litchi-xlsb/tests/theme_source.rs
)
for path in "${source_paths[@]}"; do
    [[ -f "$repo_root/$path" ]] || { echo "missing source manifest path: $path" >&2; exit 1; }
done

capture_state() {
    local destination=$1
    {
        printf 'git_head '
        git -C "$repo_root" rev-parse HEAD
        (cd "$repo_root" && sha256sum "${source_paths[@]}")
        (cd "$repo_root" && sha256sum Cargo.lock)
        if [[ -f "$repo_root/tools/perf-baseline/Cargo.lock" ]]; then
            (cd "$repo_root" && sha256sum tools/perf-baseline/Cargo.lock)
        fi
    } >"$destination"
}

state_before="$output_dir/source-state-before.txt"
state_after="$output_dir/source-state-after.txt"
capture_state "$state_before"

build_mode=release-build
if [[ -z "${PROFILE_BIN:-}" ]]; then
    build_command=(cargo "+$toolchain" build --release --locked --offline \
        --target-dir "$target_dir" --jobs "$build_jobs" -p litchi-xlsb \
        --example theme_profile)
    printf '%q ' "${build_command[@]}" >"$output_dir/release-build-command.txt"
    printf '\n' >>"$output_dir/release-build-command.txt"
    (
        cd "$repo_root"
        "${build_command[@]}"
    ) >"$output_dir/release-build.log" 2>&1
else
    build_mode=external-binary
    external_manifest=${PROFILE_BUILD_MANIFEST:-}
    [[ -n "$external_manifest" ]] || {
        echo "PROFILE_BUILD_MANIFEST is required with PROFILE_BIN" >&2
        exit 1
    }
    if [[ "$external_manifest" != /* ]]; then
        external_manifest="$repo_root/$external_manifest"
    fi
    [[ -f "$external_manifest" ]] || { echo "missing PROFILE_BUILD_MANIFEST" >&2; exit 1; }
    sha256sum "$external_manifest" >"$output_dir/external-manifest-sha256-before.txt"
fi

[[ -x "$profile_bin" ]] || { echo "profile binary is not executable: $profile_bin" >&2; exit 1; }
sha256sum "$profile_bin" >"$output_dir/binary-sha256-before.txt"

control_dir="$output_dir/generated"
control_fixture="$control_dir/theme-control-${control_bytes}.xlsb"
mkdir -p "$control_dir"
python3 "$script_dir/make-theme-control.py" \
    --base "$repo_root/test-data/poi/test-data/spreadsheet/testVarious.xlsb" \
    --output "$control_fixture" --comment-bytes "$control_bytes" \
    >"$control_dir/control-generation.txt"
python3 "$repo_root/docs/report/spec-gap-validation-evidence/xlsb-theme/verify-theme-schema.py" \
    "$control_fixture" >"$control_dir/control-schema-validation.json"

declare -a fixtures=(
    "testVarious|$repo_root/test-data/poi/test-data/spreadsheet/testVarious.xlsb"
    "62815|$repo_root/test-data/ooxml/xlsb/62815.xlsb"
    "theme-control-${control_bytes}|$control_fixture"
)
declare -a cases=(source_cold eager_cold source_warm eager_warm noop change_inverse)

for fixture_spec in "${fixtures[@]}"; do
    fixture_name=${fixture_spec%%|*}
    fixture_path=${fixture_spec#*|}
    [[ -f "$fixture_path" ]] || { echo "missing fixture: $fixture_path" >&2; exit 1; }
    for case in "${cases[@]}"; do
        for ((process=1; process<=processes; process++)); do
            suffix=""
            if (( processes > 1 )); then suffix="-p${process}"; fi
            output="$output_dir/${fixture_name}-${case}${suffix}.json"
            stderr_log="$output_dir/${fixture_name}-${case}${suffix}.stderr.log"
            "$profile_bin" --case "$case" --fixture "$fixture_path" \
                --warmup "$warmup" --samples "$samples" \
                >"$output" 2>"$stderr_log"
            python3 - "$output" "$warmup" "$samples" <<'PY'
import json
import sys

path, warmup, samples = sys.argv[1:]
report = json.loads(open(path, encoding="utf-8").read())
for gate in ("semantic_ok", "digest_stable", "preservation_ok", "inverse_ok", "change_ok", "source_observation_stable"):
    if report.get(gate) is not True:
        raise SystemExit(f"{gate} failed: {path}")
if report.get("warmup") != int(warmup) or report.get("sample_count") != int(samples):
    raise SystemExit(f"sample contract failed: {path}")
if len(report.get("samples", [])) != int(samples):
    raise SystemExit(f"raw sample cardinality failed: {path}")
PY
        done
    done
done

sha256sum "$profile_bin" >"$output_dir/binary-sha256-after.txt"
if [[ "$build_mode" == external-binary ]]; then
    sha256sum "$external_manifest" >"$output_dir/external-manifest-sha256-after.txt"
    cmp -s "$output_dir/external-manifest-sha256-before.txt" \
        "$output_dir/external-manifest-sha256-after.txt" || {
        echo "external build manifest changed during profile" >&2
        exit 1
    }
fi
capture_state "$state_after"
cmp -s "$state_before" "$state_after" || {
    echo "source, Cargo.lock, or git HEAD changed during profile" >&2
    exit 1
}
cmp -s "$output_dir/binary-sha256-before.txt" "$output_dir/binary-sha256-after.txt" || {
    echo "profile binary changed during profile" >&2
    exit 1
}

{
    printf 'build_mode %s\n' "$build_mode"
    printf 'binary_sha256_before '; awk '{print $1}' "$output_dir/binary-sha256-before.txt"
    printf 'binary_sha256_after '; awk '{print $1}' "$output_dir/binary-sha256-after.txt"
    printf 'toolchain '; rustc "+$toolchain" --version
    printf 'cargo '; cargo "+$toolchain" --version
    printf 'rustup_home %s\n' "$RUSTUP_HOME"
    printf 'target_dir %s\n' "$target_dir"
    printf 'build_jobs %s\n' "$build_jobs"
    printf 'git_commit '; git -C "$repo_root" rev-parse HEAD
    printf 'source_state_before_sha256 '; sha256sum "$state_before" | awk '{print $1}'
    printf 'source_state_after_sha256 '; sha256sum "$state_after" | awk '{print $1}'
    printf 'harness_source_sha256 '; sha256sum "$repo_root/crates/litchi-xlsb/examples/theme_profile.rs" | awk '{print $1}'
    printf 'profile_script_sha256 '; sha256sum "$script_dir/run-profile.sh" | awk '{print $1}'
    printf 'control_script_sha256 '; sha256sum "$script_dir/make-theme-control.py" | awk '{print $1}'
    printf 'schema_verifier_sha256 '; sha256sum "$repo_root/docs/report/spec-gap-validation-evidence/xlsb-theme/verify-theme-schema.py" | awk '{print $1}'
    printf 'verifier_sha256 '; sha256sum "$script_dir/verify-matrix.py" | awk '{print $1}'
    if [[ "$build_mode" == release-build ]]; then
        printf 'release_build_log_sha256 '; sha256sum "$output_dir/release-build.log" | awk '{print $1}'
    fi
    printf 'warmup %s\nsamples %s\nprocesses %s\ncontrol_bytes %s\n' "$warmup" "$samples" "$processes" "$control_bytes"
    printf 'cargo_incremental 0\n'
    for lockfile in Cargo.lock tools/perf-baseline/Cargo.lock; do
        if [[ -f "$repo_root/$lockfile" ]]; then sha256sum "$repo_root/$lockfile"; fi
    done
    for fixture_spec in "${fixtures[@]}"; do
        fixture_path=${fixture_spec#*|}
        sha256sum "$fixture_path"
    done
} >"$output_dir/provenance.txt"

python3 "$script_dir/verify-matrix.py" "$output_dir" "$warmup" "$samples" "$processes"
