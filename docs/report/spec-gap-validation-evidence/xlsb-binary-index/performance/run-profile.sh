#!/usr/bin/env bash

# Run the bounded XLSB worksheet-binary-index profile.  Each case is a fresh
# process so allocator and source counters cannot bleed between observations.
# The program reports logical in-memory ReadAt work and cache diagnostics; it
# does not measure physical I/O, RSS, or partial ZIP decompression.

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
    profile_bin="$target_dir/release/examples/binary_index_profile"
fi

[[ "$processes" =~ ^[1-9][0-9]*$ ]] || {
    echo "PROCESSES must be a positive decimal integer" >&2
    exit 1
}
[[ "$build_jobs" =~ ^[1-9][0-9]*$ ]] || {
    echo "CARGO_BUILD_JOBS must be a positive decimal integer" >&2
    exit 1
}
(( build_jobs <= 2 )) || {
    echo "CARGO_BUILD_JOBS must not exceed the bounded value 2" >&2
    exit 1
}

mkdir -p "$output_dir"

source_paths=(
    crates/litchi-xlsb/src/binary_index.rs
    crates/litchi-xlsb/src/worksheet_index.rs
    crates/litchi-xlsb/src/lib.rs
    crates/litchi-xlsb/src/workbook/source.rs
    crates/litchi-xlsb/src/workbook/mod.rs
    crates/litchi-xlsb/src/workbook/package.rs
    crates/litchi-xlsb/src/workbook/codec/package.rs
    crates/litchi-xlsb/src/cell_values/worksheet.rs
    crates/litchi-xlsb/src/writer/workbook/package.rs
    crates/litchi-xlsb/src/writer/workbook/model.rs
    crates/litchi-xlsb/src/cell_values/drawing_transfer.rs
    crates/litchi-xlsb/src/cell_values/root.rs
    crates/litchi-xlsb/src/cell_values/workbook.rs
    crates/litchi-xlsb/src/cell_watches/workbook.rs
    crates/litchi-xlsb/src/raw/kind.rs
    crates/litchi-xlsb/src/slicer/package.rs
    crates/litchi-xlsb/src/sparkline/workbook.rs
    crates/litchi-xlsb/src/timeline/package.rs
    crates/litchi-xlsb/src/xml_maps/patch.rs
    crates/litchi-xlsb/tests/worksheet_binary_index.rs
    crates/litchi-xlsb/examples/binary_index_profile.rs
)

for path in "${source_paths[@]}"; do
    [[ -f "$repo_root/$path" ]] || {
        echo "source manifest path does not exist: $path" >&2
        exit 1
    }
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
        --example binary_index_profile)
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
    [[ -f "$external_manifest" ]] || {
        echo "PROFILE_BUILD_MANIFEST does not exist: $external_manifest" >&2
        exit 1
    }
    sha256sum "$external_manifest" >"$output_dir/external-manifest-sha256-before.txt"
fi

[[ -x "$profile_bin" ]] || {
    echo "profile binary is not executable: $profile_bin" >&2
    exit 1
}

sha256sum "$profile_bin" >"$output_dir/binary-sha256-before.txt"

declare -a fixtures=(
    "sparse|test-data/poi/test-data/spreadsheet/testVarious.xlsb"
    "large|test-data/ooxml/xlsb/62815.xlsb"
)
if [[ -n "${SYNTHETIC_CELLS:-}" ]]; then
    [[ "${SYNTHETIC_CELLS}" =~ ^[1-9][0-9]*$ ]] || {
        echo "SYNTHETIC_CELLS must be a positive decimal integer" >&2
        exit 1
    }
    fixtures+=("synthetic-${SYNTHETIC_CELLS}|synthetic:${SYNTHETIC_CELLS}")
fi
declare -a cases=(indexed_cold materialize_cold indexed_warm materialize_warm)

for fixture_spec in "${fixtures[@]}"; do
    fixture_name=${fixture_spec%%|*}
    fixture_path=${fixture_spec#*|}
    if [[ "$fixture_path" == synthetic:* ]]; then
        fixture="$fixture_path"
    else
        fixture="$repo_root/$fixture_path"
        [[ -f "$fixture" ]] || { echo "missing fixture: $fixture" >&2; exit 1; }
    fi
    for case in "${cases[@]}"; do
        for ((process=1; process<=processes; process++)); do
            suffix=""
            if (( processes > 1 )); then
                suffix="-p${process}"
            fi
            output="$output_dir/${fixture_name}-${case}${suffix}.json"
            stderr_log="$output_dir/${fixture_name}-${case}${suffix}.stderr.log"
            "$profile_bin" \
                --case "$case" \
                --fixture "$fixture" \
                --warmup "$warmup" \
                --samples "$samples" \
                >"$output" 2>"$stderr_log"
            python3 - "$output" "$warmup" "$samples" <<'PY'
import json
import sys

path, warmup, samples = sys.argv[1:]
with open(path, encoding="utf-8") as stream:
    report = json.load(stream)
if report.get("semantic_ok") is not True:
    raise SystemExit(f"semantic gate failed: {path}")
if report.get("digest_stable") is not True:
    raise SystemExit(f"digest gate failed: {path}")
if report.get("source_observation_stable") is not True:
    raise SystemExit(f"source observation gate failed: {path}")
if report.get("warmup") != int(warmup) or report.get("sample_count") != int(samples):
    raise SystemExit(f"sample-count gate failed: {path}")
if len(report.get("samples", [])) != int(samples):
    raise SystemExit(f"raw sample cardinality gate failed: {path}")
PY
        done
    done
done

sha256sum "$profile_bin" >"$output_dir/binary-sha256-after.txt"
if [[ "$build_mode" == external-binary ]]; then
    sha256sum "$external_manifest" >"$output_dir/external-manifest-sha256-after.txt"
    if ! cmp -s "$output_dir/external-manifest-sha256-before.txt" \
        "$output_dir/external-manifest-sha256-after.txt"; then
        echo "external build manifest changed during profile" >&2
        exit 1
    fi
fi
capture_state "$state_after"
if ! cmp -s "$state_before" "$state_after"; then
    echo "source, Cargo.lock, or git HEAD changed during profile" >&2
    exit 1
fi
if ! cmp -s "$output_dir/binary-sha256-before.txt" \
    "$output_dir/binary-sha256-after.txt"; then
    echo "profile binary changed during profile" >&2
    exit 1
fi

{
    printf 'build_mode %s\n' "$build_mode"
    printf 'binary_sha256_before '
    awk '{print $1}' "$output_dir/binary-sha256-before.txt"
    printf 'binary_sha256_after '
    awk '{print $1}' "$output_dir/binary-sha256-after.txt"
    printf 'toolchain '
    rustc "+$toolchain" --version
    printf 'cargo '
    cargo "+$toolchain" --version
    printf 'rustup_home %s\n' "$RUSTUP_HOME"
    printf 'target_dir %s\n' "$target_dir"
    printf 'build_jobs %s\n' "$build_jobs"
    printf 'git_commit '
    git -C "$repo_root" rev-parse HEAD
    printf 'source_state_before_sha256 '
    sha256sum "$state_before" | awk '{print $1}'
    printf 'source_state_after_sha256 '
    sha256sum "$state_after" | awk '{print $1}'
    if [[ "$build_mode" == external-binary ]]; then
        printf 'external_manifest %s\n' "$external_manifest"
        printf 'external_manifest_sha256 '
        sha256sum "$external_manifest" | awk '{print $1}'
    fi
    printf 'harness_source_sha256 '
    sha256sum "$repo_root/crates/litchi-xlsb/examples/binary_index_profile.rs" | awk '{print $1}'
    printf 'profile_script_sha256 '
    sha256sum "$script_dir/run-profile.sh" | awk '{print $1}'
    if [[ "$build_mode" == release-build ]]; then
        printf 'release_build_log_sha256 '
        sha256sum "$output_dir/release-build.log" | awk '{print $1}'
    fi
    printf 'verifier_sha256 '
    sha256sum "$script_dir/verify-matrix.py" | awk '{print $1}'
    printf 'warmup %s\nsamples %s\nprocesses %s\n' "$warmup" "$samples" "$processes"
    printf 'cargo_incremental 0\n'
    for lockfile in Cargo.lock tools/perf-baseline/Cargo.lock; do
        if [[ -f "$repo_root/$lockfile" ]]; then
            sha256sum "$repo_root/$lockfile"
        fi
    done
    for fixture_spec in "${fixtures[@]}"; do
        fixture_path=${fixture_spec#*|}
        if [[ "$fixture_path" == synthetic:* ]]; then
            printf 'generated_fixture %s\n' "$fixture_path"
        else
            sha256sum "$repo_root/$fixture_path"
        fi
    done
} >"$output_dir/provenance.txt"

python3 "$script_dir/verify-matrix.py" "$output_dir" "$warmup" "$samples"
echo "wrote $((${#fixtures[@]} * ${#cases[@]} * processes)) raw profile reports to $output_dir"
