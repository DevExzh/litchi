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
HARNESS="$HERE/harness/Cargo.toml"
LOCKFILE="$HERE/harness/Cargo.lock"
SOURCE_MANIFEST="$HERE/source_manifest.py"
COMMITTED_INPUTS="$HERE/committed_inputs.py"
VERIFIER="$HERE/verify.py"
SOURCE_COMMIT=cf6fdb8e91dd232d7d762596763d2e9d8a5b9dbd
HELPER_SHA256=bec6baafcf735d778216fb54fe6299312707e912dcaed5f89de988f7112bb58e

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
git -C "$ROOT" --no-replace-objects merge-base --is-ancestor "$SOURCE_COMMIT" "$HEAD" || {
    echo "checkout does not descend from approved owner commit" >&2
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
GUARD_ARGS=(--root "$ROOT" --commit "$SOURCE_COMMIT" --helper \
    crates/litchi-pptx/tests/pptx_ink_actions.rs --helper-sha256 "$HELPER_SHA256")
for path in "${PINNED_PATHS[@]}"; do GUARD_ARGS+=(--path "$path"); done
python3 "$COMMITTED_INPUTS" "${GUARD_ARGS[@]}" > /dev/null

mkdir -p -- "$(dirname -- "$RESULTS")" "$(dirname -- "$TARGET")"
mkdir -- "$RESULTS"
mkdir -- "$TARGET"
cleanup() {
    status=$?
    if [[ "$status" -eq 0 ]]; then
        find "$TARGET" -depth -delete
    else
        echo "retaining failed profile target for diagnosis: $TARGET" >&2
    fi
    exit "$status"
}
trap cleanup EXIT

export CARGO_TARGET_DIR="$TARGET"
export CARGO_INCREMENTAL=0
export LC_ALL=C
unset RUSTFLAGS CARGO_ENCODED_RUSTFLAGS RUSTC_BOOTSTRAP RUSTDOCFLAGS

COMMANDS="$RESULTS/commands.jsonl"
METADATA_BEFORE="$RESULTS/metadata-before.json"
METADATA_AFTER="$RESULTS/metadata-after.json"
MANIFEST_BEFORE="$RESULTS/source-manifest-before.txt"
MANIFEST_AFTER="$RESULTS/source-manifest-after.txt"
GIT_STATUS_BEFORE="$RESULTS/git-status-before.txt"
GIT_STATUS_AFTER="$RESULTS/git-status-after.txt"
HOST="$RESULTS/host.txt"
BUILD_LOG="$RESULTS/build.log"
BIN="$TARGET/release/pptx-ink-actions-performance"
: > "$COMMANDS"

record_command() {
    local label=$1
    shift
    python3 - "$COMMANDS" "$label" "$ROOT" "$@" <<'PY'
import json
import os
import pathlib
import sys

output, label, cwd, *argv = sys.argv[1:]
with pathlib.Path(output).open("a") as stream:
    stream.write(json.dumps({
        "label": label,
        "cwd": cwd,
        "argv": argv,
        "environment": {
            "CARGO_TARGET_DIR": os.environ.get("CARGO_TARGET_DIR"),
            "CARGO_INCREMENTAL": os.environ.get("CARGO_INCREMENTAL"),
            "LC_ALL": os.environ.get("LC_ALL"),
            "RUSTFLAGS": None,
            "CARGO_ENCODED_RUSTFLAGS": None,
            "RUSTC_BOOTSTRAP": None,
            "RUSTDOCFLAGS": None,
        },
    }, sort_keys=True) + "\n")
PY
}

printf '%s\n' \
    "source_commit=$SOURCE_COMMIT" \
    "git_head=$HEAD" \
    "processes=$PROCESSES" \
    "warmup=$WARMUP" \
    "samples=$SAMPLES" \
    "rustc=$(rustc -vV | tr '\n' ' ')" \
    "cargo=$(cargo -V)" \
    "uname=$(uname -a)" \
    "cpu_model=$(awk -F: '/model name|Hardware/ {gsub(/^ +/, \"\", $2); print $2; exit}' /proc/cpuinfo 2>/dev/null || true)" \
    "memory=$(awk '/MemTotal:/ {print $2 \" \" $3; exit}' /proc/meminfo 2>/dev/null || true)" \
    > "$HOST"
git -C "$ROOT" status --porcelain --untracked-files=all -- \
    crates/litchi-pptx crates/litchi-opc crates/litchi-drawingml > "$GIT_STATUS_BEFORE"

record_command cargo-metadata-before cargo metadata --format-version=1 --locked --offline --manifest-path "$HARNESS"
cargo metadata --format-version=1 --locked --offline --manifest-path "$HARNESS" > "$METADATA_BEFORE"

EXTRAS=(
    "$ROOT/Cargo.toml" "$ROOT/rust-toolchain.toml" "$ROOT/.cargo/config.toml"
    "$HARNESS" "$LOCKFILE" "$HERE/harness/main.rs" "$HERE/harness/adapter.rs"
    "$HERE/harness/support.rs" "$SOURCE_MANIFEST" "$COMMITTED_INPUTS"
    "$VERIFIER" "$HERE/run_profile.sh" "$HERE/README.md" "$HERE/PLAN.md"
    "$HERE/report.md" "$HERE/requirements.md" "$HERE/corpus-manifest.json"
    "$HERE/root-review.md"
)
MANIFEST_ARGS=(python3 "$SOURCE_MANIFEST" --metadata "$METADATA_BEFORE" --root "$ROOT"
    --output "$MANIFEST_BEFORE" --git-commit "$HEAD")
for extra in "${EXTRAS[@]}"; do MANIFEST_ARGS+=(--extra "$extra"); done
record_command source-manifest-before "${MANIFEST_ARGS[@]}"
"${MANIFEST_ARGS[@]}"

record_command cargo-build-release cargo build --release --locked --offline --manifest-path "$HARNESS"
cargo build --release --locked --offline --manifest-path "$HARNESS" > "$BUILD_LOG" 2>&1
[[ -x "$BIN" ]] || { echo "release binary missing: $BIN" >&2; exit 1; }
sha256sum "$BIN" > "$RESULTS/binary.sha256"
{
    printf 'binary=%s\n' "$BIN"
    cat "$RESULTS/binary.sha256"
    printf 'source_commit=%s\n' "$SOURCE_COMMIT"
    printf 'git_head=%s\n' "$HEAD"
    printf 'rustc -vV:\n'; rustc -vV
    printf 'cargo=%s\n' "$(cargo -V)"
    printf 'target=%s\n' "$TARGET"
    printf 'allocator=CountingAllocator (process-local GlobalAlloc observer)\n'
    printf 'rss=/usr/bin/time -v Maximum resident set size (kbytes)\n'
    printf 'timing_scope=absolute latency, allocation receipts, and process RSS; no speedup claim\n'
} > "$RESULTS/build-provenance.txt"

record_command host-probe "$BIN" --host-probe
"$BIN" --host-probe > "$RESULTS/host-probe.json" 2> "$RESULTS/host-probe.stderr"
record_command correctness-matrix "$BIN" --matrix-correctness
"$BIN" --matrix-correctness > "$RESULTS/matrix-correctness.json" 2> "$RESULTS/matrix-correctness.stderr"
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
        record_command "lane-${lane}-p${process}" /usr/bin/time -v -o "$timing" "$BIN" \
            --lane "$lane" --warmup "$WARMUP" --samples "$SAMPLES"
        /usr/bin/time -v -o "$timing" "$BIN" --lane "$lane" \
            --warmup "$WARMUP" --samples "$SAMPLES" > "$report" 2> "$stderr"
    done
done

sha256sum "$BIN" > "$RESULTS/binary-after.sha256"
cmp -s "$RESULTS/binary.sha256" "$RESULTS/binary-after.sha256"
record_command cargo-metadata-after cargo metadata --format-version=1 --locked --offline --manifest-path "$HARNESS"
cargo metadata --format-version=1 --locked --offline --manifest-path "$HARNESS" > "$METADATA_AFTER"
MANIFEST_ARGS_AFTER=(python3 "$SOURCE_MANIFEST" --metadata "$METADATA_AFTER" --root "$ROOT"
    --output "$MANIFEST_AFTER" --git-commit "$HEAD")
for extra in "${EXTRAS[@]}"; do MANIFEST_ARGS_AFTER+=(--extra "$extra"); done
record_command source-manifest-after "${MANIFEST_ARGS_AFTER[@]}"
"${MANIFEST_ARGS_AFTER[@]}"
cmp -s "$MANIFEST_BEFORE" "$MANIFEST_AFTER"
git -C "$ROOT" status --porcelain --untracked-files=all -- \
    crates/litchi-pptx crates/litchi-opc crates/litchi-drawingml > "$GIT_STATUS_AFTER"
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
} > "$RESULTS/source-provenance.txt"

record_command verify python3 "$VERIFIER" --root "$ROOT" --evidence "$HERE" --results "$RESULTS" \
    --manifest "$MANIFEST_BEFORE" --manifest-after "$MANIFEST_AFTER" \
    --metadata-before "$METADATA_BEFORE" --metadata-after "$METADATA_AFTER" \
    --corpus "$HERE/corpus-manifest.json" --output "$RESULTS/verification.json"
python3 "$VERIFIER" --root "$ROOT" --evidence "$HERE" --results "$RESULTS" \
    --manifest "$MANIFEST_BEFORE" --manifest-after "$MANIFEST_AFTER" \
    --metadata-before "$METADATA_BEFORE" --metadata-after "$METADATA_AFTER" \
    --corpus "$HERE/corpus-manifest.json" --output "$RESULTS/verification.json"
echo "PPTX InkAction profile verified: $RESULTS"
