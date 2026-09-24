#!/usr/bin/env bash
set -euo pipefail

if [[ "${PROFILE_FROZEN:-0}" != 1 ]]; then
    echo "profile is intentionally gated: set PROFILE_FROZEN=1 only after the approved InkAction source freeze" >&2
    exit 2
fi

HERE=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
ROOT=$(cd -- "$HERE/../../../../" && pwd)
HARNESS="$HERE/harness/Cargo.toml"
MANIFEST_TOOL="$HERE/source_manifest.py"
PROFILE_PINS="$HERE/profile_pins.py"
RESULTS=$(realpath -m -- "${PROFILE_RESULTS_DIR:-$HERE/results}")
REPORT_OUTPUT=$(realpath -m -- "${PROFILE_REPORT_OUTPUT:-$HERE/report.md}")
TARGET_INPUT=${CARGO_TARGET_DIR:-/var/tmp/litchi-ink-action-edit-profile-target}
TARGET=$(realpath -m -- "$TARGET_INPUT")
WARMUP=${WARMUP:-2}
SAMPLES=${SAMPLES:-20}
PROCESSES=${PROCESSES:-3}
PROFILE_ARM=${PROFILE_ARM:-baseline}
CURRENT_COMMIT=$(git -C "$ROOT" rev-parse HEAD)

LANES=(
    draft_small_8 draft_scaled_128 draft_near_1024 draft_opaque_64
    scalar_edit_small_8 scalar_edit_scaled_128 scalar_edit_near_1024
    scalar_batch_scaled_128 scalar_batch_near_1024
    scalar_coalesce_scaled_128 scalar_coalesce_near_1024
    no_op_small_8 no_op_scaled_128 no_op_near_1024
    add_small_8 add_scaled_128 add_near_1024
    insert_batch_scaled_128 insert_batch_near_1024
    remove_small_8 remove_scaled_128 remove_near_1024
    remove_batch_scaled_128 remove_batch_near_1024
    clear_batch_scaled_128 clear_batch_near_1024
    move_small_8 move_scaled_128 move_near_1024
    move_batch_scaled_128 move_batch_near_1024
    cap_refusal_small_8 cap_refusal_scaled_128 cap_refusal_near_1024
)

if [[ ! -f "$PROFILE_PINS" ]]; then
    echo "profile source pin manifest is missing: $PROFILE_PINS" >&2
    exit 1
fi
PROFILE_SOURCE_PIN=""
APPROVED_BASE_COMMIT=""
declare -A EXPECTED_SOURCE_HASHES=()
while IFS=$'\t' read -r kind relative digest; do
    case "$kind" in
        base)
            if [[ -n "$APPROVED_BASE_COMMIT" || -z "$relative" || -n "$digest" ]]; then
                echo "malformed profile base pin manifest for arm $PROFILE_ARM" >&2
                exit 2
            fi
            APPROVED_BASE_COMMIT="$relative"
            ;;
        pin)
            if [[ -n "$PROFILE_SOURCE_PIN" || -z "$relative" || -n "$digest" ]]; then
                echo "malformed profile source pin manifest for arm $PROFILE_ARM" >&2
                exit 2
            fi
            PROFILE_SOURCE_PIN="$relative"
            ;;
        source)
            if [[ -z "$relative" || -z "$digest" || -n "${EXPECTED_SOURCE_HASHES[$ROOT/$relative]:-}" ]]; then
                echo "malformed profile source hash manifest for arm $PROFILE_ARM" >&2
                exit 2
            fi
            EXPECTED_SOURCE_HASHES["$ROOT/$relative"]="$digest"
            ;;
        "")
            ;;
        *)
            echo "unknown profile source pin manifest record: $kind" >&2
            exit 2
            ;;
    esac
done < <(python3 "$PROFILE_PINS" --arm "$PROFILE_ARM")
if [[ -z "$APPROVED_BASE_COMMIT" || -z "$PROFILE_SOURCE_PIN" || "${#EXPECTED_SOURCE_HASHES[@]}" -ne 5 ]]; then
    echo "incomplete profile source pin manifest for arm $PROFILE_ARM" >&2
    exit 2
fi

for variable in RUSTFLAGS CARGO_ENCODED_RUSTFLAGS RUSTC_BOOTSTRAP RUSTDOCFLAGS; do
    if [[ -n "${!variable:-}" ]]; then
        echo "refusing profile with ${variable} set" >&2
        exit 2
    fi
done
while IFS='=' read -r variable _; do
    if [[ "$variable" == CARGO_PROFILE_RELEASE_* ]]; then
        echo "refusing profile with ${variable} set" >&2
        exit 2
    fi
done < <(env)
if [[ ! -x /usr/bin/time ]]; then
    echo "refusing profile because /usr/bin/time -v is unavailable" >&2
    exit 2
fi
if ! git -C "$ROOT" merge-base --is-ancestor "$APPROVED_BASE_COMMIT" "$CURRENT_COMMIT"; then
    echo "refusing profile without the approved InkAction source commit in history" >&2
    exit 2
fi
if ! git -C "$ROOT" merge-base --is-ancestor "$PROFILE_SOURCE_PIN" "$CURRENT_COMMIT"; then
    echo "refusing profile without the selected $PROFILE_ARM source pin in history: $PROFILE_SOURCE_PIN" >&2
    exit 2
fi
relevant_status=$(git -C "$ROOT" status --short -- \
    crates/litchi-drawingml/src/ink crates/litchi-drawingml/tests/ink_action_edit.rs \
    crates/litchi-drawingml/tests/ink_action_id_boundaries.rs)
if [[ -n "$relevant_status" ]]; then
    echo "refusing profile with relevant source changes:" >&2
    printf '%s\n' "$relevant_status" >&2
    exit 2
fi
for source in "${!EXPECTED_SOURCE_HASHES[@]}"; do
    if [[ ! -f "$source" ]]; then
        echo "approved source input is missing: $source" >&2
        exit 1
    fi
    actual=$(sha256sum "$source" | cut -d' ' -f1)
    if [[ "$actual" != "${EXPECTED_SOURCE_HASHES[$source]}" ]]; then
        echo "approved source hash changed: $source" >&2
        exit 1
    fi
done
if [[ ! -f "$HARNESS" || ! -f "$MANIFEST_TOOL" ]]; then
    echo "profile harness input is missing" >&2
    exit 1
fi

case "$PROCESSES:$WARMUP:$SAMPLES" in
    *[!0-9:]*|*:0:*|*:*:0*)
        echo "PROCESSES, WARMUP, and SAMPLES must be positive integers" >&2
        exit 2
        ;;
esac
if [[ "$PROCESSES" -ne 3 || "$WARMUP" -lt 2 || "$SAMPLES" -lt 20 ]]; then
    echo "acceptance requires 3 processes, at least 2 warm-ups, and at least 20 samples" >&2
    exit 2
fi
if [[ "$TARGET" == "/" || "$TARGET" == "$ROOT" || "$TARGET" == "$ROOT/target" || "$TARGET" == "$HERE" ]]; then
    echo "refusing unsafe Cargo target path: $TARGET" >&2
    exit 2
fi
if [[ -e "$TARGET" && "${ALLOW_EXISTING_TARGET:-0}" != 1 ]]; then
    echo "refusing to reuse an existing target: $TARGET" >&2
    exit 2
fi

TARGET_CREATED=0
if [[ ! -e "$TARGET" ]]; then
    TARGET_CREATED=1
fi
remove_tree() {
    if [[ -e "$1" ]]; then
        find "$1" -depth -delete
    fi
}
cleanup() {
    status=$?
    if [[ "$TARGET_CREATED" == 1 ]]; then
        remove_tree "$TARGET"
    fi
    exit "$status"
}
trap cleanup EXIT

mkdir -p "$RESULTS"
# Remove only this runner's named acceptance receipts. Review notes, exploratory
# outputs, and receipts for other lanes remain available for audit.
for lane in "${LANES[@]}"; do
    for process in 1 2 3; do
        rm -f "$RESULTS/${lane}-p${process}.json" \
            "$RESULTS/${lane}-p${process}.time.txt" \
            "$RESULTS/${lane}-p${process}.stderr.log"
    done
done
rm -f "$RESULTS/metadata-before.json" "$RESULTS/metadata-after.json" \
    "$RESULTS/build.log" "$RESULTS/commands.txt" "$RESULTS/host.txt" \
    "$RESULTS/source-manifest-before.txt" "$RESULTS/source-manifest-after.txt" \
    "$RESULTS/source-provenance.txt" "$RESULTS/binary.sha256" \
    "$RESULTS/binary-after.sha256" "$RESULTS/build-provenance.txt" \
    "$RESULTS/verification.json" "$REPORT_OUTPUT"

export CARGO_TARGET_DIR="$TARGET"
export CARGO_INCREMENTAL=0
export LC_ALL=C
unset RUSTFLAGS CARGO_ENCODED_RUSTFLAGS RUSTC_BOOTSTRAP RUSTDOCFLAGS

{
    printf '%s\n' "root=$ROOT"
    printf '%s\n' "harness=$HARNESS"
    printf '%s\n' "target=$TARGET"
    printf '%s\n' "profile_arm=$PROFILE_ARM"
    printf '%s\n' "source_pin=$PROFILE_SOURCE_PIN"
    printf '%s\n' "warmup=$WARMUP"
    printf '%s\n' "samples_per_process=$SAMPLES"
    printf '%s\n' "fresh_processes_per_lane=$PROCESSES"
    printf '%s\n' "lanes=${LANES[*]}"
    printf '%s\n' "command=cargo metadata --format-version=1 --locked --offline --manifest-path $HARNESS"
    printf '%s\n' "command=cargo build --release --locked --offline --manifest-path $HARNESS"
    printf '%s\n' "command=/usr/bin/time -v $TARGET/release/ink-action-edit-profile --lane {lane} --warmup $WARMUP --samples $SAMPLES"
} >"$RESULTS/commands.txt"
{
    date -u '+utc=%Y-%m-%dT%H:%M:%SZ'
    rustc -vV
    cargo -V
    uname -a
    if command -v lscpu >/dev/null 2>&1; then lscpu | sed -n '1,32p'; fi
} >"$RESULTS/host.txt"

metadata_args=(--metadata "$RESULTS/metadata-before.json" --root "$ROOT" --output "$RESULTS/source-manifest-before.txt" --git-commit "$CURRENT_COMMIT")
for extra in "$ROOT/Cargo.toml" "$HARNESS" "$HERE/harness/Cargo.lock" \
    "$MANIFEST_TOOL" "$PROFILE_PINS" "$HERE/run_profile.sh" "$HERE/smoke.sh" "$HERE/summarize.py" \
    "$HERE/verify.py" "$HERE/test_verify.py" "$HERE/test_smoke_target.py" \
    "$HERE/test_source_snapshot.py" \
    "$HERE/README.md" "$HERE/requirements.md"; do
    metadata_args+=(--extra "$extra")
done
if [[ -f "$ROOT/rust-toolchain.toml" ]]; then
    metadata_args+=(--extra "$ROOT/rust-toolchain.toml")
fi

cargo metadata --format-version=1 --locked --offline --manifest-path "$HARNESS" \
    >"$RESULTS/metadata-before.json"
python3 "$MANIFEST_TOOL" "${metadata_args[@]}"

cargo build --release --locked --offline --manifest-path "$HARNESS" \
    >"$RESULTS/build.log" 2>&1
BIN="$TARGET/release/ink-action-edit-profile"
if [[ ! -x "$BIN" ]]; then
    echo "profile binary was not produced: $BIN" >&2
    exit 1
fi
sha256sum "$BIN" >"$RESULTS/binary.sha256"
{
    printf 'binary=%s\n' "$BIN"
    cat "$RESULTS/binary.sha256"
    printf 'profile_arm=%s\n' "$PROFILE_ARM"
    printf 'source_pin=%s\n' "$PROFILE_SOURCE_PIN"
    printf '%s\n' 'rustc -vV:'
    rustc -vV
    printf '%s\n' "cargo=$(cargo -V)"
    printf '%s\n' "target=$TARGET"
    printf '%s\n' "cargo_incremental=$CARGO_INCREMENTAL"
    printf '%s\n' 'flags=none'
    printf '%s\n' 'allocator=CountingAllocator (process-local GlobalAlloc observer)'
} >"$RESULTS/build-provenance.txt"

for lane in "${LANES[@]}"; do
    for process in 1 2 3; do
        report="$RESULTS/${lane}-p${process}.json"
        timing="$RESULTS/${lane}-p${process}.time.txt"
        printf '%s\n' \
            "run=/usr/bin/time -v $BIN --lane $lane --warmup $WARMUP --samples $SAMPLES (fresh_process=$process)" \
            >>"$RESULTS/commands.txt"
        /usr/bin/time -v -o "$timing" "$BIN" \
            --lane "$lane" --warmup "$WARMUP" --samples "$SAMPLES" \
            >"$report" 2>"$RESULTS/${lane}-p${process}.stderr.log"
    done
done

sha256sum "$BIN" >"$RESULTS/binary-after.sha256"
cmp -s "$RESULTS/binary.sha256" "$RESULTS/binary-after.sha256"

cargo metadata --format-version=1 --locked --offline --manifest-path "$HARNESS" \
    >"$RESULTS/metadata-after.json"
metadata_args_after=(--metadata "$RESULTS/metadata-after.json" --root "$ROOT" --output "$RESULTS/source-manifest-after.txt" --git-commit "$CURRENT_COMMIT")
for extra in "$ROOT/Cargo.toml" "$HARNESS" "$HERE/harness/Cargo.lock" \
    "$MANIFEST_TOOL" "$PROFILE_PINS" "$HERE/run_profile.sh" "$HERE/smoke.sh" "$HERE/summarize.py" \
    "$HERE/verify.py" "$HERE/test_verify.py" "$HERE/test_smoke_target.py" \
    "$HERE/test_source_snapshot.py" \
    "$HERE/README.md" "$HERE/requirements.md"; do
    metadata_args_after+=(--extra "$extra")
done
if [[ -f "$ROOT/rust-toolchain.toml" ]]; then
    metadata_args_after+=(--extra "$ROOT/rust-toolchain.toml")
fi
python3 "$MANIFEST_TOOL" "${metadata_args_after[@]}"
cmp -s "$RESULTS/source-manifest-before.txt" "$RESULTS/source-manifest-after.txt"

{
    printf 'profile_arm=%s\n' "$PROFILE_ARM"
    printf 'source_pin=%s\n' "$PROFILE_SOURCE_PIN"
    printf 'source_manifest_before_sha256='
    sha256sum "$RESULTS/source-manifest-before.txt" | cut -d' ' -f1
    printf 'source_manifest_after_sha256='
    sha256sum "$RESULTS/source-manifest-after.txt" | cut -d' ' -f1
    printf 'approved_base_commit=%s\n' "$APPROVED_BASE_COMMIT"
    printf 'git_head=%s\n' "$CURRENT_COMMIT"
    printf 'git_status_relevant=%s\n' "$relevant_status"
    printf 'profile_pins_sha256='
    sha256sum "$PROFILE_PINS" | cut -d' ' -f1
    printf '%s\n' 'source_sha256:'
    for source in "${!EXPECTED_SOURCE_HASHES[@]}"; do sha256sum "$source"; done | sort -k2
    printf '%s\n' 'harness_sha256:'
    find "$HERE/harness" -type f -print0 | sort -z | xargs -0 sha256sum
    printf '%s\n' 'test_sha256:'
    sha256sum "$HERE/test_verify.py" "$HERE/test_smoke_target.py" "$HERE/test_source_snapshot.py"
} >"$RESULTS/source-provenance.txt"

python3 "$HERE/summarize.py" --results "$RESULTS" --output "$REPORT_OUTPUT"
python3 "$HERE/verify.py" --results "$RESULTS" --report "$REPORT_OUTPUT"
