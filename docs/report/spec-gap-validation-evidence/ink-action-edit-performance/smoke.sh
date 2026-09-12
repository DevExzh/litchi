#!/usr/bin/env bash

set -euo pipefail

if [ "${PROFILE_FROZEN:-0}" != 1 ]; then
    echo "smoke is intentionally gated: set PROFILE_FROZEN=1 for the approved source" >&2
    exit 2
fi

HERE=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
ROOT=$(CDPATH= cd -- "$HERE/../../../../" && pwd)
HARNESS="$HERE/harness/Cargo.toml"
PROFILE_PINS="$HERE/profile_pins.py"
RESULTS="$HERE/results/smoke"
TARGET=${CARGO_TARGET_DIR:-/var/tmp/litchi-ink-action-edit-smoke-target}
PROFILE_ARM=${PROFILE_ARM:-baseline}
if [ ! -f "$PROFILE_PINS" ]; then
    echo "profile source pin manifest is missing: $PROFILE_PINS" >&2
    exit 1
fi
APPROVED_BASE_COMMIT=""
PROFILE_SOURCE_PIN=""
declare -A EXPECTED_SOURCE_HASHES=()
while IFS=$'\t' read -r kind relative digest; do
    case "$kind" in
        base)
            if [ -n "$APPROVED_BASE_COMMIT" ] || [ -z "$relative" ] || [ -n "$digest" ]; then
                echo "malformed profile base pin manifest for arm $PROFILE_ARM" >&2
                exit 2
            fi
            APPROVED_BASE_COMMIT="$relative"
            ;;
        pin)
            if [ -n "$PROFILE_SOURCE_PIN" ] || [ -z "$relative" ] || [ -n "$digest" ]; then
                echo "malformed profile source pin manifest for arm $PROFILE_ARM" >&2
                exit 2
            fi
            PROFILE_SOURCE_PIN="$relative"
            ;;
        source)
            if [ -z "$relative" ] || [ -z "$digest" ] || [ -n "${EXPECTED_SOURCE_HASHES[$ROOT/$relative]:-}" ]; then
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
if [ -z "$APPROVED_BASE_COMMIT" ] || [ -z "$PROFILE_SOURCE_PIN" ] || [ "${#EXPECTED_SOURCE_HASHES[@]}" -ne 5 ]; then
    echo "incomplete profile source pin manifest for arm $PROFILE_ARM" >&2
    exit 2
fi
EXPECTED_ACTIONS_EDIT_SHA256=${EXPECTED_SOURCE_HASHES["$ROOT/crates/litchi-drawingml/src/ink/actions_edit.rs"]}
LANES="draft_opaque_64 scalar_edit_scaled_128 scalar_batch_scaled_128 scalar_batch_near_1024 scalar_coalesce_scaled_128 scalar_coalesce_near_1024 no_op_small_8 add_small_8 insert_batch_scaled_128 insert_batch_near_1024 remove_batch_scaled_128 remove_batch_near_1024 clear_batch_scaled_128 clear_batch_near_1024 move_batch_scaled_128 move_batch_near_1024 cap_refusal_small_8"

SOURCE_FILES="
$ROOT/crates/litchi-drawingml/src/ink/mod.rs
$ROOT/crates/litchi-drawingml/src/ink/actions.rs
$ROOT/crates/litchi-drawingml/src/ink/actions_edit.rs
$ROOT/crates/litchi-drawingml/tests/ink_action_edit.rs
$ROOT/crates/litchi-drawingml/tests/ink_action_id_boundaries.rs"

CURRENT_COMMIT=$(git -C "$ROOT" rev-parse HEAD)
if ! git -C "$ROOT" merge-base --is-ancestor "$APPROVED_BASE_COMMIT" "$CURRENT_COMMIT"; then
    echo "approved InkAction smoke source commit is not in history" >&2
    exit 1
fi
if ! git -C "$ROOT" merge-base --is-ancestor "$PROFILE_SOURCE_PIN" "$CURRENT_COMMIT"; then
    echo "selected $PROFILE_ARM smoke source pin is not in history: $PROFILE_SOURCE_PIN" >&2
    exit 1
fi
for variable in RUSTFLAGS CARGO_ENCODED_RUSTFLAGS RUSTC_BOOTSTRAP RUSTDOCFLAGS; do
    if [ -n "${!variable:-}" ]; then
        echo "refusing smoke with ${variable} set" >&2
        exit 2
    fi
done
for path in $SOURCE_FILES; do
    if [ ! -f "$path" ]; then
        echo "ink-action smoke source input is missing: $path" >&2
        exit 1
    fi
done
for path in "${!EXPECTED_SOURCE_HASHES[@]}"; do
    actual=$(sha256sum "$path" | cut -d' ' -f1)
    if [ "$actual" != "${EXPECTED_SOURCE_HASHES[$path]}" ]; then
        echo "approved source hash changed: $path" >&2
        exit 1
    fi
done

mkdir -p "$RESULTS"
TARGET=$(realpath -m -- "$TARGET")
TARGET_OWNED=0
if [ ! -e "$TARGET" ]; then
    TARGET_OWNED=1
fi
case "$TARGET" in
    /|"$ROOT"|"$ROOT/target"|"$HERE")
        echo "refusing unsafe smoke target: $TARGET" >&2
        exit 2
        ;;
esac
if [ "$TARGET_OWNED" -eq 0 ] && [ "${ALLOW_EXISTING_TARGET:-0}" != 1 ]; then
    echo "refusing to reuse an existing smoke target: $TARGET" >&2
    exit 2
fi
TEMP_DIR=$(mktemp -d "${TMPDIR:-/tmp}/litchi-ink-action-edit-smoke.XXXXXX")

remove_tree() {
    if [ -e "$1" ]; then
        find "$1" -depth -delete
    fi
}

write_source_inputs() {
    {
        printf 'base_commit='
        git -C "$ROOT" rev-parse HEAD
        printf 'profile_arm=%s\n' "$PROFILE_ARM"
        printf 'source_pin=%s\n' "$PROFILE_SOURCE_PIN"
        printf 'approved_base_commit=%s\n' "$APPROVED_BASE_COMMIT"
        printf 'source_sha256:\n'
        for path in $SOURCE_FILES; do
            sha256sum "$path"
        done
    } >"$1"
}

cleanup() {
    status=$?
    remove_tree "$TEMP_DIR"
    if [ "${CLEAN_TARGET:-1}" = 1 ] && [ "$TARGET_OWNED" -eq 1 ]; then
        remove_tree "$TARGET"
    fi
    exit "$status"
}
trap cleanup EXIT

for lane in $LANES; do
    rm -f "$RESULTS/${lane}.json"
done
rm -f "$RESULTS/source-inputs-before.txt" "$RESULTS/source-inputs-after.txt" \
    "$RESULTS/source-provenance.txt" "$RESULTS/build.log" "$RESULTS/commands.txt" \
    "$RESULTS/binary.sha256" "$RESULTS/binary-after.sha256" "$RESULTS/build-provenance.txt"

{
    printf '%s\n' "root=$ROOT"
    printf '%s\n' "harness=$HARNESS"
    printf '%s\n' "target=$TARGET"
    printf '%s\n' "profile_arm=$PROFILE_ARM"
    printf '%s\n' "source_pin=$PROFILE_SOURCE_PIN"
    printf '%s\n' 'warmup=1'
    printf '%s\n' 'samples_per_lane=1'
    printf '%s\n' "lanes=$LANES"
    printf '%s\n' "approved_actions_edit_sha256=$EXPECTED_ACTIONS_EDIT_SHA256"
    printf '%s\n' "approved_commit=$PROFILE_SOURCE_PIN"
} >"$RESULTS/commands.txt"

export CARGO_TARGET_DIR="$TARGET"
export CARGO_INCREMENTAL=0
write_source_inputs "$RESULTS/source-inputs-before.txt"

cargo build --release --locked --offline --manifest-path "$HARNESS" >"$RESULTS/build.log" 2>&1
BIN="$TARGET/release/ink-action-edit-profile"
if [ ! -x "$BIN" ]; then
    echo "smoke binary was not produced: $BIN" >&2
    exit 1
fi
sha256sum "$BIN" >"$RESULTS/binary.sha256"
{
    printf 'binary=%s\n' "$BIN"
    cat "$RESULTS/binary.sha256"
    printf 'profile_arm=%s\n' "$PROFILE_ARM"
    printf 'source_pin=%s\n' "$PROFILE_SOURCE_PIN"
    printf '%s\n' 'flags=none'
    printf '%s\n' 'cargo_incremental=0'
} >"$RESULTS/build-provenance.txt"

for lane in $LANES; do
    report="$RESULTS/${lane}.json"
    printf '%s\n' "run=$BIN --lane $lane --warmup 1 --samples 1" >>"$RESULTS/commands.txt"
    "$BIN" --lane "$lane" --warmup 1 --samples 1 >"$report"
done

write_source_inputs "$RESULTS/source-inputs-after.txt"
cmp -s "$RESULTS/source-inputs-before.txt" "$RESULTS/source-inputs-after.txt"
sha256sum "$BIN" >"$RESULTS/binary-after.sha256"
cmp -s "$RESULTS/binary.sha256" "$RESULTS/binary-after.sha256"

{
    printf 'base_commit='
    git -C "$ROOT" rev-parse HEAD
    printf 'profile_arm=%s\n' "$PROFILE_ARM"
    printf 'source_pin=%s\n' "$PROFILE_SOURCE_PIN"
    printf 'approved_base_commit=%s\n' "$APPROVED_BASE_COMMIT"
    printf 'source_inputs_before_sha256='
    sha256sum "$RESULTS/source-inputs-before.txt" | cut -d' ' -f1
    printf 'source_inputs_after_sha256='
    sha256sum "$RESULTS/source-inputs-after.txt" | cut -d' ' -f1
    printf 'approved_actions_edit_sha256=%s\n' "$EXPECTED_ACTIONS_EDIT_SHA256"
    printf 'profile_pins_sha256='
    sha256sum "$PROFILE_PINS" | cut -d' ' -f1
    printf 'git_status_relevant='
    git -C "$ROOT" status --short -- \
        crates/litchi-drawingml/src/ink crates/litchi-drawingml/tests/ink_action_edit.rs \
        crates/litchi-drawingml/tests/ink_action_id_boundaries.rs | tr '\n' '|'
    printf '\nsource_sha256:\n'
    for path in $SOURCE_FILES; do
        sha256sum "$path"
    done
    printf '%s\n' 'harness_sha256:'
    find "$HERE/harness" -type f -print0 | sort -z | xargs -0 sha256sum
    printf '%s\n' 'test_sha256:'
    sha256sum "$HERE/test_verify.py" "$HERE/test_smoke_target.py" "$HERE/test_source_snapshot.py"
} >"$RESULTS/source-provenance.txt"

python3 - "$RESULTS" "$LANES" <<'PY'
import json
import sys
from pathlib import Path

results = Path(sys.argv[1])
lanes = sys.argv[2].split()
batch_prefixes = (
    "scalar_batch_",
    "scalar_coalesce_",
    "insert_batch_",
    "remove_batch_",
    "clear_batch_",
    "move_batch_",
)
for lane in lanes:
    path = results / f"{lane}.json"
    value = json.loads(path.read_text())
    assert value["lane"] == lane, path
    assert value["schema"] == "ink-action-edit-profile-v2", path
    assert type(value["pid"]) is int and value["pid"] > 0, path
    assert value["warmup"] == 1 and value["sample_count"] == 1, path
    sample = value["samples"]
    assert len(sample) == 1, path
    sample = sample[0]
    assert sample["opaque_preserved"] is True, path
    assert sample["alloc_balance_ok"] is True, path
    assert sample["alloc_invalid"] is False and sample["alloc_failed"] == 0, path
    if value["lane"].startswith("cap_refusal_"):
        assert sample["actual_success"] is False, path
        assert sample["semantic_ok"] is None, path
        assert sample["source_exact"] is None, path
        assert sample["source_shared"] is None, path
        assert sample["inverse_ok"] is None, path
        assert sample["output_exact"] is None, path
        assert sample["rejection_ok"] is True, path
        assert sample["rejection_source_unchanged"] is True, path
        assert sample["rejection_state_unchanged"] is True, path
        assert sample["rejection_resource"] == "ink action output bytes", path
        assert sample["rejection_limit"] == value["input_bytes"], path
        assert sample["rejection_pre_source_hash"] == sample["rejection_post_source_hash"], path
    elif value["lane"].startswith("draft_"):
        assert sample["actual_success"] is True, path
        assert sample["semantic_ok"] is True, path
        assert sample["source_exact"] is None, path
        assert sample["source_shared"] is None, path
        assert sample["inverse_ok"] is None, path
        assert sample["output_exact"] is True, path
        assert sample["rejection_ok"] is None, path
        assert sample["rejection_source_unchanged"] is None, path
        assert sample["rejection_state_unchanged"] is None, path
        assert sample["rejection_resource"] is None and sample["rejection_limit"] is None, path
        assert sample["rejection_pre_source_hash"] is None and sample["rejection_post_source_hash"] is None, path
    else:
        assert sample["actual_success"] is True, path
        assert sample["semantic_ok"] is True, path
        assert sample["source_exact"] is True, path
        assert sample["inverse_ok"] is True, path
        assert sample["output_exact"] is True, path
        assert sample["source_shared"] is (value["lane"].startswith("no_op_")), path
        assert sample["rejection_ok"] is None, path
        assert sample["rejection_source_unchanged"] is None, path
        assert sample["rejection_state_unchanged"] is None, path
        assert sample["rejection_resource"] is None and sample["rejection_limit"] is None, path
        assert sample["rejection_pre_source_hash"] is None and sample["rejection_post_source_hash"] is None, path
    if value["lane"].startswith(batch_prefixes):
        assert value["operation_count"] > 1, path
print(f"validated {len(lanes)} bounded smoke lanes")
PY
