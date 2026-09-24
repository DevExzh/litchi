#!/usr/bin/env bash
set -euo pipefail

[[ "${PPTX_INK_ACTIONS_PROFILE_FROZEN:-0}" == 1 ]] || {
    echo "profile is frozen; set PPTX_INK_ACTIONS_PROFILE_FROZEN=1 after review" >&2
    exit 2
}
[[ "${PPTX_INK_ACTIONS_API_WIRED:-0}" == 1 ]] || {
    echo "profile API wiring is not authorized" >&2
    exit 2
}
[[ "${PPTX_INK_ACTIONS_ISOLATED_SOURCE:-0}" == 1 ]] || {
    echo "profile requires an explicitly isolated source checkout" >&2
    exit 2
}
[[ "${PPTX_INK_ACTIONS_TIMING_AUTHORIZED:-0}" == 1 ]] || {
    echo "timing capture is not authorized" >&2
    exit 2
}

for variable in RUSTFLAGS CARGO_ENCODED_RUSTFLAGS RUSTC_BOOTSTRAP RUSTDOCFLAGS; do
    [[ -z "${!variable:-}" ]] || {
        echo "profile refuses inherited ${variable}" >&2
        exit 2
    }
done
[[ -x /usr/bin/time ]] || {
    echo "profile requires /usr/bin/time -v" >&2
    exit 2
}

HERE=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd -P)
ROOT=$(cd -- "$HERE/../../../.." && pwd -P)
cd -- "$ROOT"
HARNESS="$HERE/harness/Cargo.toml"
LOCKFILE="$HERE/harness/Cargo.lock"
SOURCE_MANIFEST="$HERE/source_manifest.py"
COMMITTED_INPUTS="$HERE/committed_inputs.py"
VERIFIER="$HERE/verify.py"
CLEANUP_HELPER="$HERE/cleanup_target.sh"
SEMANTIC_OWNER_COMMIT=cf6fdb8e91dd232d7d762596763d2e9d8a5b9dbd
PRODUCTION_SOURCE_BASELINE_COMMIT=2a2ffa1cae4e6b7070082768ce84483e5d411dc8
SEMANTIC_OWNER_DESIGN_PATH=docs/report/spec-gap-validation-evidence/pptx-ink-actions-design.md
SEMANTIC_OWNER_DESIGN_SHA256=30b78cca84c4ca24ae44f3d3694c3097f54b5e5a1f2004af9ce5007bcaf4173d
SEMANTIC_OWNER_DESIGN_GIT_BLOB=597400950b1027c47cd6e4cbbedd23915bc0980e
HELPER_SHA256=bec6baafcf735d778216fb54fe6299312707e912dcaed5f89de988f7112bb58e

[[ -x "$CLEANUP_HELPER" ]] || {
    echo "profile cleanup helper is missing or not executable" >&2
    exit 2
}

RESULTS_INPUT=${PPTX_INK_ACTIONS_PROFILE_RESULTS:-}
TARGET_INPUT=${PPTX_INK_ACTIONS_PROFILE_TARGET:-/var/tmp/litchi-pptx-ink-actions-profile-target}
[[ -n "$RESULTS_INPUT" ]] || {
    echo "set PPTX_INK_ACTIONS_PROFILE_RESULTS to a fresh output directory" >&2
    exit 2
}
RESULTS=$(realpath -m -- "$RESULTS_INPUT")
TARGET=$(realpath -m -- "$TARGET_INPUT")
case "$RESULTS/" in "$ROOT/"*) echo "results must be outside checkout" >&2; exit 2;; esac
case "$TARGET/" in "$ROOT/"*) echo "target must be outside checkout" >&2; exit 2;; esac
[[ "$RESULTS" != "/" && "$TARGET" != "/" && "$RESULTS" != "$TARGET" ]] || {
    echo "results and target paths are unsafe or overlap" >&2
    exit 2
}
case "$RESULTS/" in "$TARGET/"*) echo "results and target overlap" >&2; exit 2;; esac
case "$TARGET/" in "$RESULTS/"*) echo "target and results overlap" >&2; exit 2;; esac
[[ ! -e "$RESULTS" && ! -L "$RESULTS" && ! -e "$TARGET" && ! -L "$TARGET" ]] || {
    echo "results and target must be fresh absent paths" >&2
    exit 2
}

PROCESSES=${PPTX_INK_ACTIONS_PROCESSES:-3}
WARMUP=${PPTX_INK_ACTIONS_WARMUP:-2}
SAMPLES=${PPTX_INK_ACTIONS_SAMPLES:-20}
[[ "$PROCESSES" == 3 && "$WARMUP" == 2 && "$SAMPLES" == 20 ]] || {
    echo "profile requires exactly 3 processes, 2 warmups, and 20 samples" >&2
    exit 2
}

HEAD=$(git -C "$ROOT" --no-replace-objects rev-parse HEAD)
git -C "$ROOT" --no-replace-objects merge-base --is-ancestor "$SEMANTIC_OWNER_COMMIT" "$HEAD" || {
    echo "checkout does not descend from semantic owner commit" >&2
    exit 2
}
git -C "$ROOT" --no-replace-objects merge-base --is-ancestor "$PRODUCTION_SOURCE_BASELINE_COMMIT" "$HEAD" || {
    echo "checkout does not descend from production source baseline" >&2
    exit 2
}
[[ -z "$(git -C "$ROOT" --no-replace-objects status --porcelain --untracked-files=all)" ]] || {
    echo "profile requires a clean isolated checkout" >&2
    exit 2
}
[[ -f "$HARNESS" && -f "$LOCKFILE" ]] || {
    echo "isolated harness Cargo.toml/Cargo.lock must be materialized before build" >&2
    exit 2
}
git -C "$ROOT" ls-files --error-unmatch -- "${LOCKFILE#"$ROOT/"}" >/dev/null || {
    echo "isolated harness Cargo.lock is not committed" >&2
    exit 2
}

mkdir -p -- "$(dirname -- "$RESULTS")"
mkdir -- "$RESULTS"
COMMANDS="$RESULTS/commands.jsonl"
: > "$COMMANDS"
export CARGO_TARGET_DIR="$TARGET"
export CARGO_INCREMENTAL=0
export LC_ALL=C
export PPTX_INK_ACTIONS_CAPTURE_HEAD="$HEAD"
unset RUSTFLAGS CARGO_ENCODED_RUSTFLAGS RUSTC_BOOTSTRAP RUSTDOCFLAGS

record_command_start() {
    local label=$1
    local stdout_path=$2
    local stderr_path=$3
    shift
    shift
    shift
    python3 - "$COMMANDS" "$label" "$ROOT" "$$" "$stdout_path" "$stderr_path" "$@" <<'PY'
import json
import os
import pathlib
import sys
import time

output, label, cwd, pid, stdout_path, stderr_path, *argv = sys.argv[1:]
with pathlib.Path(output).open("a") as stream:
    stream.write(json.dumps({
        "event": "start",
        "label": label,
        "cwd": cwd,
        "pid": int(pid),
        "argv": argv,
        "outputs": {
            "stdout": stdout_path or None,
            "stderr": stderr_path or None,
        },
        "started_unix_ns": time.time_ns(),
        "environment": {
            "CARGO_TARGET_DIR": os.environ.get("CARGO_TARGET_DIR"),
            "CARGO_INCREMENTAL": os.environ.get("CARGO_INCREMENTAL"),
            "LC_ALL": os.environ.get("LC_ALL"),
            "PPTX_INK_ACTIONS_CAPTURE_HEAD": os.environ.get("PPTX_INK_ACTIONS_CAPTURE_HEAD"),
            "RUSTFLAGS": None,
            "CARGO_ENCODED_RUSTFLAGS": None,
            "RUSTC_BOOTSTRAP": None,
            "RUSTDOCFLAGS": None,
        },
    }, sort_keys=True) + "\n")
PY
}

record_command_exit() {
    local label=$1
    local status=$2
    local stdout_path=$3
    local stderr_path=$4
    python3 - "$COMMANDS" "$label" "$ROOT" "$$" "$status" "$stdout_path" "$stderr_path" <<'PY'
import json
import os
import pathlib
import sys
import time

output, label, cwd, pid, status, stdout_path, stderr_path = sys.argv[1:]
with pathlib.Path(output).open("a") as stream:
    stream.write(json.dumps({
        "event": "exit",
        "label": label,
        "cwd": cwd,
        "pid": int(pid),
        "exit_code": int(status),
        "outputs": {
            "stdout": stdout_path or None,
            "stderr": stderr_path or None,
        },
        "finished_unix_ns": time.time_ns(),
        "environment": {
            "CARGO_TARGET_DIR": os.environ.get("CARGO_TARGET_DIR"),
            "CARGO_INCREMENTAL": os.environ.get("CARGO_INCREMENTAL"),
            "LC_ALL": os.environ.get("LC_ALL"),
            "PPTX_INK_ACTIONS_CAPTURE_HEAD": os.environ.get("PPTX_INK_ACTIONS_CAPTURE_HEAD"),
            "RUSTFLAGS": None,
            "CARGO_ENCODED_RUSTFLAGS": None,
            "RUSTC_BOOTSTRAP": None,
            "RUSTDOCFLAGS": None,
        },
    }, sort_keys=True) + "\n")
PY
}

run_recorded() {
    local label=$1
    shift
    local stdout_path=""
    local stderr_path=""
    while [[ "$#" -gt 0 ]]; do
        case "$1" in
            --stdout)
                [[ "$#" -ge 2 ]] || { echo "run_recorded --stdout needs a path" >&2; return 2; }
                stdout_path=$2
                shift 2
                ;;
            --stderr)
                [[ "$#" -ge 2 ]] || { echo "run_recorded --stderr needs a path" >&2; return 2; }
                stderr_path=$2
                shift 2
                ;;
            *)
                break
                ;;
        esac
    done
    [[ "$#" -gt 0 ]] || { echo "run_recorded has no command: $label" >&2; return 2; }
    record_command_start "$label" "$stdout_path" "$stderr_path" "$@"
    set +e
    if [[ -n "$stdout_path" && -n "$stderr_path" ]]; then
        "$@" > "$stdout_path" 2> "$stderr_path"
    elif [[ -n "$stdout_path" ]]; then
        "$@" > "$stdout_path"
    elif [[ -n "$stderr_path" ]]; then
        "$@" 2> "$stderr_path"
    else
        "$@"
    fi
    local status=$?
    set -e
    record_command_exit "$label" "$status" "$stdout_path" "$stderr_path"
    return "$status"
}

PINNED_PATHS=(
    crates/litchi-pptx/src/lib.rs
    crates/litchi-pptx/src/package/codec.rs
    crates/litchi-pptx/src/package/model.rs
    crates/litchi-pptx/src/presentation/model.rs
    crates/litchi-pptx/src/presentation/package.rs
    crates/litchi-pptx/src/presentation/embedded/ink_actions/model.rs
    crates/litchi-pptx/src/presentation/embedded/ink_actions/codec.rs
    crates/litchi-pptx/src/presentation/embedded/ink_actions/package.rs
    crates/litchi-pptx/src/presentation/embedded/ink_actions/transaction.rs
    crates/litchi-pptx/tests/pptx_ink_actions.rs
)
GUARD_ARGS=(--root "$ROOT" --semantic-owner-commit "$SEMANTIC_OWNER_COMMIT" \
    --production-source-baseline-commit "$PRODUCTION_SOURCE_BASELINE_COMMIT" --helper \
    crates/litchi-pptx/tests/pptx_ink_actions.rs --helper-sha256 "$HELPER_SHA256" \
    --manifest-path "$HARNESS" --exclude-package litchi-pptx-ink-actions-performance)
for path in "${PINNED_PATHS[@]}"; do GUARD_ARGS+=(--path "$path"); done
GUARD_ARGS+=(--semantic-owner-extra "$ROOT/$SEMANTIC_OWNER_DESIGN_PATH" \
    --semantic-owner-extra-sha256 "$SEMANTIC_OWNER_DESIGN_SHA256" \
    --semantic-owner-extra-blob "$SEMANTIC_OWNER_DESIGN_GIT_BLOB")
PRODUCTION_WORKSPACE_EXTRAS=(
    "$ROOT/Cargo.toml" "$ROOT/rust-toolchain.toml" "$ROOT/.cargo/config.toml"
    "$ROOT/rustfmt.toml" "$ROOT/clippy.toml" "$ROOT/deny.toml"
)
for path in "${PRODUCTION_WORKSPACE_EXTRAS[@]}"; do GUARD_ARGS+=(--production-extra "$path"); done
run_recorded committed-inputs-guard --stdout "$RESULTS/committed-inputs-guard.stdout" \
    --stderr "$RESULTS/committed-inputs-guard.stderr" \
    python3 "$COMMITTED_INPUTS" "${GUARD_ARGS[@]}"

mkdir -p -- "$(dirname -- "$TARGET")"
mkdir -- "$TARGET"
TARGET_SENTINEL="$TARGET/.pptx-ink-actions-profile-target.sentinel"
SUCCESS_SENTINEL="$TARGET/.pptx-ink-actions-profile-success.sentinel"
cleanup() {
    status=$?
    "$CLEANUP_HELPER" "$status" "$ROOT" "$TARGET" "$TARGET_SENTINEL" || status=$?
    exit "$status"
}
trap cleanup EXIT

printf 'pptx-ink-actions-profile-target-v1\nroot=%s\ntarget=%s\n' "$ROOT" "$TARGET" > "$TARGET_SENTINEL"
chmod 600 -- "$TARGET_SENTINEL"

METADATA_BEFORE="$RESULTS/metadata-before.json"
METADATA_AFTER="$RESULTS/metadata-after.json"
MANIFEST_BEFORE="$RESULTS/source-manifest-before.txt"
MANIFEST_AFTER="$RESULTS/source-manifest-after.txt"
GIT_STATUS_BEFORE="$RESULTS/git-status-before.txt"
GIT_STATUS_AFTER="$RESULTS/git-status-after.txt"
HOST="$RESULTS/host.txt"
BUILD_LOG="$RESULTS/build.log"
BIN="$TARGET/release/pptx-ink-actions-performance"

printf '%s\n' \
    "source_commit=$SEMANTIC_OWNER_COMMIT" \
    "semantic_owner_commit=$SEMANTIC_OWNER_COMMIT" \
    "production_source_baseline_commit=$PRODUCTION_SOURCE_BASELINE_COMMIT" \
    "semantic_owner_design_path=$SEMANTIC_OWNER_DESIGN_PATH" \
    "semantic_owner_design_sha256=$SEMANTIC_OWNER_DESIGN_SHA256" \
    "semantic_owner_design_git_blob=$SEMANTIC_OWNER_DESIGN_GIT_BLOB" \
    "capture_head=$HEAD" \
    "git_head=$HEAD" \
    "processes=$PROCESSES" \
    "warmup=$WARMUP" \
    "samples=$SAMPLES" \
    "rustc=$(rustc -vV | tr '\n' ' ')" \
    "cargo=$(cargo -V)" \
    "uname=$(uname -a)" \
    "hostname=$(hostname)" \
    "loadavg=$(cat /proc/loadavg 2>/dev/null || true)" \
    "cpu_model=$(awk -F: '/model name|Hardware/ {gsub(/^ +/, \"\", $2); print $2; exit}' /proc/cpuinfo 2>/dev/null || true)" \
    "memory=$(awk '/MemTotal:/ {print $2 \" \" $3; exit}' /proc/meminfo 2>/dev/null || true)" \
    > "$HOST"
git -C "$ROOT" status --porcelain --untracked-files=all > "$GIT_STATUS_BEFORE"

run_recorded cargo-metadata-before --stdout "$METADATA_BEFORE" \
    --stderr "$RESULTS/cargo-metadata-before.stderr" \
    cargo metadata --format-version=1 --locked --offline --manifest-path "$HARNESS"

CONTEXT_EXTRAS=(
    # ADRs are current capture context, hashed against capture HEAD. The
    # separately pinned design document is the semantic-owner authority.
    "$ROOT/docs/adr/0001-priorities-and-api-layers.md"
    "$ROOT/docs/adr/0003-snapshots-edits-and-patches.md"
    "$ROOT/docs/adr/0005-io-memory-and-performance.md"
    "$ROOT/docs/adr/0006-validation-security-and-compatibility.md"
)
EXTRAS=(
    "$ROOT/Cargo.toml" "$ROOT/rust-toolchain.toml" "$ROOT/.cargo/config.toml"
    "$ROOT/rustfmt.toml" "$ROOT/clippy.toml" "$ROOT/deny.toml"
    "$HARNESS" "$LOCKFILE" "$HERE/harness/main.rs" "$HERE/harness/adapter.rs"
    "$HERE/harness/support.rs" "$SOURCE_MANIFEST" "$COMMITTED_INPUTS"
    "$VERIFIER" "$CLEANUP_HELPER" "$HERE/run_profile.sh" "$HERE/README.md" "$HERE/PLAN.md"
    "$HERE/report.md" "$HERE/requirements.md" "$HERE/corpus-manifest.json"
    "$HERE/source-contract.json" "$HERE/scaffold_tests.py"
    "$HERE/root-review.md"
)
MANIFEST_ARGS=(python3 "$SOURCE_MANIFEST" --metadata "$METADATA_BEFORE" --root "$ROOT"
    --output "$MANIFEST_BEFORE" --git-commit "$HEAD"
    --semantic-owner-commit "$SEMANTIC_OWNER_COMMIT"
    --production-source-baseline-commit "$PRODUCTION_SOURCE_BASELINE_COMMIT"
    --production-exclude-package litchi-pptx-ink-actions-performance)
MANIFEST_ARGS+=(--semantic-owner-extra "$ROOT/$SEMANTIC_OWNER_DESIGN_PATH"
    --semantic-owner-extra-sha256 "$SEMANTIC_OWNER_DESIGN_SHA256"
    --semantic-owner-extra-blob "$SEMANTIC_OWNER_DESIGN_GIT_BLOB")
for extra in "${EXTRAS[@]}"; do MANIFEST_ARGS+=(--extra "$extra"); done
for extra in "${CONTEXT_EXTRAS[@]}"; do MANIFEST_ARGS+=(--context-extra "$extra"); done
for extra in "${PRODUCTION_WORKSPACE_EXTRAS[@]}"; do MANIFEST_ARGS+=(--production-extra "$extra"); done
run_recorded source-manifest-before --stdout "$RESULTS/source-manifest-before.stdout" \
    --stderr "$RESULTS/source-manifest-before.stderr" "${MANIFEST_ARGS[@]}"

run_recorded cargo-build-release --stdout "$BUILD_LOG" --stderr "$RESULTS/build.stderr" \
    cargo build --release --locked --offline --manifest-path "$HARNESS"
[[ -x "$BIN" ]] || { echo "release binary missing: $BIN" >&2; exit 1; }
sha256sum "$BIN" > "$RESULTS/binary.sha256"
{
    printf 'binary=%s\n' "$BIN"
    cat "$RESULTS/binary.sha256"
    printf 'source_commit=%s\n' "$SEMANTIC_OWNER_COMMIT"
    printf 'semantic_owner_commit=%s\n' "$SEMANTIC_OWNER_COMMIT"
    printf 'production_source_baseline_commit=%s\n' "$PRODUCTION_SOURCE_BASELINE_COMMIT"
    printf 'semantic_owner_design_path=%s\n' "$SEMANTIC_OWNER_DESIGN_PATH"
    printf 'semantic_owner_design_sha256=%s\n' "$SEMANTIC_OWNER_DESIGN_SHA256"
    printf 'semantic_owner_design_git_blob=%s\n' "$SEMANTIC_OWNER_DESIGN_GIT_BLOB"
    printf 'capture_head=%s\n' "$HEAD"
    printf 'git_head=%s\n' "$HEAD"
    printf 'rustc -vV:\n'; rustc -vV
    printf 'cargo=%s\n' "$(cargo -V)"
    printf 'target=%s\n' "$TARGET"
    printf 'allocator=CountingAllocator (process-local GlobalAlloc observer)\n'
    printf 'rss=/usr/bin/time -v Maximum resident set size (kbytes)\n'
    printf 'timing_scope=absolute latency, allocation receipts, and process RSS; no speedup claim\n'
} > "$RESULTS/build-provenance.txt"

run_recorded host-probe --stdout "$RESULTS/host-probe.json" --stderr "$RESULTS/host-probe.stderr" \
    "$BIN" --host-probe
run_recorded correctness-matrix --stdout "$RESULTS/matrix-correctness.json" \
    --stderr "$RESULTS/matrix-correctness.stderr" "$BIN" --matrix-correctness
[[ ! -s "$RESULTS/host-probe.stderr" && ! -s "$RESULTS/matrix-correctness.stderr" ]] || {
    echo "correctness preflight wrote stderr" >&2
    exit 1
}

LANES=(
    package_read_tiny_shared presentation_read_small_shared package_read_medium_shared
    package_read_small_distinct package_read_large_shared package_read_large_distinct
    package_read_near_shared package_read_near_distinct package_read_multislide_shared
    package_read_multislide_distinct package_scalar_edit_small_shared
    presentation_scalar_edit_small_shared package_noop_small_shared
    presentation_noop_small_shared package_apply_medium_shared
    package_apply_medium_distinct package_inverse_small_shared
    package_inverse_case_equivalent package_save_medium_shared
    package_save_reopen_medium_shared stale_owner stale_owner_rels stale_target
    stale_content_type signed_noop signed_changed_refusal opaque_mce_scalar_edit
    opaque_mce_save_reopen unknown_outbound_read_edit strict_shared_edit
    limit_anchor_one_under limit_anchor_exact limit_anchor_one_over
    limit_target_one_under limit_target_exact limit_target_one_over
    limit_aggregate_one_under limit_aggregate_exact limit_aggregate_one_over
    limit_graph_one_under limit_graph_exact limit_graph_one_over
)

for lane in "${LANES[@]}"; do
    for process in 1 2 3; do
        report="$RESULTS/${lane}-p${process}.json"
        timing="$RESULTS/${lane}-p${process}.time.txt"
        stderr="$RESULTS/${lane}-p${process}.stderr.log"
        run_recorded "lane-${lane}-p${process}" --stdout "$report" --stderr "$stderr" \
            /usr/bin/time -v -o "$timing" "$BIN" --lane "$lane" \
            --warmup "$WARMUP" --samples "$SAMPLES"
    done
done

sha256sum "$BIN" > "$RESULTS/binary-after.sha256"
cmp -s "$RESULTS/binary.sha256" "$RESULTS/binary-after.sha256"
run_recorded cargo-metadata-after --stdout "$METADATA_AFTER" \
    --stderr "$RESULTS/cargo-metadata-after.stderr" \
    cargo metadata --format-version=1 --locked --offline --manifest-path "$HARNESS"
MANIFEST_ARGS_AFTER=(python3 "$SOURCE_MANIFEST" --metadata "$METADATA_AFTER" --root "$ROOT"
    --output "$MANIFEST_AFTER" --git-commit "$HEAD"
    --semantic-owner-commit "$SEMANTIC_OWNER_COMMIT"
    --production-source-baseline-commit "$PRODUCTION_SOURCE_BASELINE_COMMIT"
    --production-exclude-package litchi-pptx-ink-actions-performance)
MANIFEST_ARGS_AFTER+=(--semantic-owner-extra "$ROOT/$SEMANTIC_OWNER_DESIGN_PATH"
    --semantic-owner-extra-sha256 "$SEMANTIC_OWNER_DESIGN_SHA256"
    --semantic-owner-extra-blob "$SEMANTIC_OWNER_DESIGN_GIT_BLOB")
for extra in "${EXTRAS[@]}"; do MANIFEST_ARGS_AFTER+=(--extra "$extra"); done
for extra in "${CONTEXT_EXTRAS[@]}"; do MANIFEST_ARGS_AFTER+=(--context-extra "$extra"); done
for extra in "${PRODUCTION_WORKSPACE_EXTRAS[@]}"; do MANIFEST_ARGS_AFTER+=(--production-extra "$extra"); done
run_recorded source-manifest-after --stdout "$RESULTS/source-manifest-after.stdout" \
    --stderr "$RESULTS/source-manifest-after.stderr" "${MANIFEST_ARGS_AFTER[@]}"
cmp -s "$MANIFEST_BEFORE" "$MANIFEST_AFTER"
git -C "$ROOT" status --porcelain --untracked-files=all > "$GIT_STATUS_AFTER"
{
    printf 'source_manifest_sha256='
    sha256sum "$MANIFEST_BEFORE" | awk '{print $1}'
    printf 'cargo_lock_sha256='
    sha256sum "$LOCKFILE" | awk '{print $1}'
    printf 'corpus_manifest_sha256='
    sha256sum "$HERE/corpus-manifest.json" | awk '{print $1}'
    printf 'helper_sha256='
    sha256sum "$ROOT/crates/litchi-pptx/tests/pptx_ink_actions.rs" | awk '{print $1}'
    printf 'host_sha256='
    sha256sum "$HOST" | awk '{print $1}'
    printf 'semantic_owner_commit=%s\n' "$SEMANTIC_OWNER_COMMIT"
    printf 'production_source_baseline_commit=%s\n' "$PRODUCTION_SOURCE_BASELINE_COMMIT"
    printf 'semantic_owner_design_path=%s\n' "$SEMANTIC_OWNER_DESIGN_PATH"
    printf 'semantic_owner_design_sha256=%s\n' "$SEMANTIC_OWNER_DESIGN_SHA256"
    printf 'semantic_owner_design_git_blob=%s\n' "$SEMANTIC_OWNER_DESIGN_GIT_BLOB"
    printf 'capture_head=%s\n' "$HEAD"
} > "$RESULTS/source-provenance.txt"

run_recorded verify --stdout "$RESULTS/verify.stdout" --stderr "$RESULTS/verify.stderr" \
    python3 "$VERIFIER" --root "$ROOT" --evidence "$HERE" --results "$RESULTS" \
    --manifest "$MANIFEST_BEFORE" --manifest-after "$MANIFEST_AFTER" \
    --metadata-before "$METADATA_BEFORE" --metadata-after "$METADATA_AFTER" \
    --corpus "$HERE/corpus-manifest.json" --output "$RESULTS/verification.json"
printf 'pptx-ink-actions-profile-success-v1\nroot=%s\ntarget=%s\n' "$ROOT" "$TARGET" > "$SUCCESS_SENTINEL"
chmod 600 -- "$SUCCESS_SENTINEL"
echo "PPTX InkAction profile verified: $RESULTS"
