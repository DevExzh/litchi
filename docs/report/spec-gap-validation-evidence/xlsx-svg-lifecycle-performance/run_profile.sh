#!/usr/bin/env bash
set -euo pipefail

if [[ "${PROFILE_FROZEN:-}" != "1" ]]; then
    echo "refusing to profile an unfrozen XLSX SVG lifecycle owner; set PROFILE_FROZEN=1" >&2
    exit 2
fi
if [[ "${XLSX_SVG_PROFILE_API_WIRED:-}" != "1" ]]; then
    echo "refusing to profile before the XLSX SVG API adapter is review-wired; set XLSX_SVG_PROFILE_API_WIRED=1 only after semantic review and freeze" >&2
    exit 2
fi
for variable in RUSTFLAGS CARGO_ENCODED_RUSTFLAGS RUSTC_BOOTSTRAP; do
    if [[ -n "${!variable:-}" ]]; then
        echo "refusing profile with ${variable} set" >&2
        exit 2
    fi
done
if [[ ! -x /usr/bin/time ]]; then
    echo "refusing profile because /usr/bin/time -v is unavailable for RSS evidence" >&2
    exit 2
fi

HERE=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd -P)
ROOT=$(cd -- "$HERE/../../../.." && pwd -P)
HARNESS="$HERE/harness/Cargo.toml"
if [[ -z "${XLSX_SVG_PROFILE_RESULTS:-}" ]]; then
    echo "set XLSX_SVG_PROFILE_RESULTS to a fresh output directory outside the checkout" >&2
    exit 2
fi
RESULTS=$(realpath -m -- "$XLSX_SVG_PROFILE_RESULTS")
TARGET_INPUT=${CARGO_TARGET_DIR:-/var/tmp/litchi-xlsx-svg-lifecycle-profile-target}
if [[ "$TARGET_INPUT" = /* ]]; then
    TARGET=$(realpath -m -- "$TARGET_INPUT")
else
    TARGET=$(realpath -m -- "$PWD/$TARGET_INPUT")
fi
MANIFEST_TOOL="$HERE/source_manifest.py"
PROFILE_PINS="$HERE/profile_pins.py"
COMMITTED_INPUTS="$HERE/committed_inputs.py"
LANES=(
    capture_native_fixture
    capture_raster_two_cell_small capture_raster_two_cell_large
    capture_raster_one_cell_small capture_raster_one_cell_large
    capture_raster_absolute_small capture_raster_absolute_large
    capture_attached_two_cell_small capture_attached_two_cell_large
    capture_attached_one_cell_small capture_attached_one_cell_large
    capture_attached_absolute_small capture_attached_absolute_large
    clone_raster_small clone_raster_large
    clone_attached_small clone_attached_large
    clone_captured_owner_small clone_captured_owner_large
    inventory_shared_256 inventory_shared_1024
    inventory_distinct_256 inventory_distinct_1024
    namespace_heavy namespace_limit_refusal
    attach_end_to_end_two_cell_small attach_end_to_end_two_cell_large
    attach_end_to_end_one_cell_small attach_end_to_end_one_cell_large
    attach_end_to_end_absolute_small attach_end_to_end_absolute_large
    strict_attach_end_to_end_two_cell_small
    inverse_attach_detach_two_cell_small inverse_attach_detach_two_cell_large
    inverse_attach_detach_one_cell_small inverse_attach_detach_one_cell_large
    inverse_attach_detach_absolute_small inverse_attach_detach_absolute_large
    detach_end_to_end_shared_first_two_cell
    detach_end_to_end_shared_first_one_cell
    detach_end_to_end_shared_first_absolute
    detach_end_to_end_shared_final_two_cell
    detach_end_to_end_shared_final_one_cell
    detach_end_to_end_shared_final_absolute
    strict_detach_end_to_end_shared_final_two_cell
    incoming_edge_shared_final_two_cell
    detach_end_to_end_distinct_two_cell_small
    detach_end_to_end_distinct_two_cell_large
    detach_end_to_end_distinct_one_cell_small
    detach_end_to_end_distinct_one_cell_large
    detach_end_to_end_distinct_absolute_small
    detach_end_to_end_distinct_absolute_large
    same_picture_attach_detach_two_cell
    same_picture_attach_detach_one_cell
    same_picture_attach_detach_absolute
    multisheet_attach_detach
    noop_detach_two_cell noop_detach_one_cell noop_detach_absolute
    limit_small limit_large
    mixed_caps_rejection
    malformed_duplicate_owner malformed_mce_owner
    malformed_linked_owner malformed_unknown_uri
    multi_picture_same_drawing_16 multi_picture_same_drawing_64
    multi_picture_same_drawing_256
)
PROCESSES=${PROCESSES:-3}
WARMUP=${WARMUP:-2}
SAMPLES=${SAMPLES:-20}

case "$PROCESSES:$WARMUP:$SAMPLES" in
    *[!0-9:]*|0:*|*:0:*|*:*:0*)
        echo "PROCESSES, WARMUP, and SAMPLES must be positive integers" >&2
        exit 2
        ;;
esac
if [[ "$PROCESSES" -ne 3 || "$WARMUP" -lt 2 || "$SAMPLES" -lt 20 ]]; then
    echo "acceptance requires 3 processes, at least 2 warmups, and at least 20 samples" >&2
    exit 2
fi
# Keep retained evidence disjoint from both source and disposable build output.
if [[ "$RESULTS" == "$ROOT" || "$RESULTS" == "$ROOT/"* || "$ROOT" == "$RESULTS/"* ]]; then
    echo "profile results must be outside the checkout: $RESULTS" >&2
    exit 2
fi
if [[ "$TARGET" == "$ROOT" || "$TARGET" == "$ROOT/"* || "$ROOT" == "$TARGET/"* || "$TARGET" == "/" ]]; then
    echo "Cargo target must be outside the checkout: $TARGET" >&2
    exit 2
fi
if [[ "$RESULTS" == "/" || "$RESULTS" == "$TARGET" || "$RESULTS" == "$TARGET/"* || "$TARGET" == "$RESULTS/"* ]]; then
    echo "profile results and Cargo target must be disjoint: $RESULTS / $TARGET" >&2
    exit 2
fi
if [[ -e "$RESULTS" || -L "$RESULTS" || -L "$XLSX_SVG_PROFILE_RESULTS" ]]; then
    echo "refusing existing profile output; choose a fresh directory: $RESULTS" >&2
    exit 2
fi
if [[ -e "$TARGET" || -L "$TARGET" ]]; then
    if [[ "${ALLOW_EXISTING_TARGET:-}" != "1" || ! -d "$TARGET" ]]; then
        echo "refusing to reuse an existing target; choose a fresh target or set ALLOW_EXISTING_TARGET=1" >&2
        exit 2
    fi
fi

TARGET_CREATED=0
cleanup_target() {
    if [[ "$TARGET_CREATED" == "1" ]]; then
        find "$TARGET" -depth -delete
    fi
}
trap cleanup_target EXIT

# mkdir without -p atomically refuses an output directory created after preflight.
# Failed-run diagnostics are retained here; no prior receipt is ever removed.
mkdir -p -- "$(dirname -- "$RESULTS")" "$(dirname -- "$TARGET")"
mkdir -- "$RESULTS"
if [[ ! -e "$TARGET" ]]; then
    mkdir -- "$TARGET"
    TARGET_CREATED=1
fi

unset RUSTFLAGS CARGO_ENCODED_RUSTFLAGS RUSTC_BOOTSTRAP
export CARGO_TARGET_DIR="$TARGET"
export CARGO_INCREMENTAL=0
export LC_ALL=C

cargo metadata --format-version=1 --locked --offline --manifest-path "$HARNESS" \
    >"$RESULTS/metadata-before.json"
CURRENT_COMMIT=$(git -C "$ROOT" rev-parse HEAD)
SOURCE_PIN=$(python3 "$PROFILE_PINS" | sed -n 's/^source_pin=//p')
if [[ -z "$SOURCE_PIN" ]]; then
    echo "profile source pin manifest did not produce a source pin" >&2
    exit 1
fi
if [[ -n "${XLSX_SVG_PROFILE_SOURCE_PIN:-}" && "$XLSX_SVG_PROFILE_SOURCE_PIN" != "$SOURCE_PIN" ]]; then
    echo "requested XLSX SVG source pin differs from the approved profile pin" >&2
    exit 2
fi
if ! git -C "$ROOT" merge-base --is-ancestor "$SOURCE_PIN" "$CURRENT_COMMIT"; then
    echo "refusing profile without the approved XLSX SVG source commit in history: $SOURCE_PIN" >&2
    exit 2
fi
mapfile -t PINNED_SOURCES < <(python3 "$PROFILE_PINS" --paths)
guard_args=(--root "$ROOT" --commit "$SOURCE_PIN")
for source in "${PINNED_SOURCES[@]}"; do
    guard_args+=(--path "$source")
done
python3 "$COMMITTED_INPUTS" "${guard_args[@]}"
python3 "$MANIFEST_TOOL" \
    --metadata "$RESULTS/metadata-before.json" \
    --root "$ROOT" \
    --output "$RESULTS/source-manifest-before.txt" \
    --extra "$ROOT/Cargo.toml" \
    --extra "$ROOT/rust-toolchain.toml" \
    --extra "$ROOT/.cargo/config.toml" \
    --extra "$HARNESS" \
    --extra "$HERE/harness/Cargo.lock" \
    --extra "$HERE/harness/adapter.rs" \
    --extra "$HERE/harness/support.rs" \
    --extra "$MANIFEST_TOOL" \
    --extra "$PROFILE_PINS" \
    --extra "$COMMITTED_INPUTS" \
    --extra "$HERE/run_profile.sh" \
    --extra "$HERE/summarize.py" \
    --extra "$HERE/uncertainty.py" \
    --extra "$HERE/verify.py" \
    --extra "$HERE/README.md" \
    --extra "$HERE/requirements.md" \
    --extra "$HERE/root-review.md" \
    --extra "$HERE/corpus-manifest.json" \
    --extra "$HERE/fixtures/tdf169496_hidden_graphic.xlsx" \
    --extra "$ROOT/docs/adr/0001-priorities-and-api-layers.md" \
    --extra "$ROOT/docs/adr/0005-io-memory-and-performance.md" \
    --extra "$ROOT/docs/report/spec-gap-validation-evidence/xlsx-svg-lifecycle-design.md" \
    --git-commit "$CURRENT_COMMIT"

cargo build --release --locked --offline --manifest-path "$HARNESS" \
    >"$RESULTS/build.log" 2>&1
BIN="$TARGET/release/xlsx-svg-lifecycle-profile"
if [[ ! -x "$BIN" ]]; then
    echo "profile binary was not produced: $BIN" >&2
    exit 1
fi
sha256sum "$BIN" >"$RESULTS/binary.sha256"
{
    printf 'binary=%s\n' "$BIN"
    cat "$RESULTS/binary.sha256"
    printf '%s\n' 'rustc -vV:'
    rustc -vV
    printf '%s\n' "cargo=$(cargo -V)"
    printf '%s\n' "target=$TARGET"
    printf '%s\n' "cargo_incremental=$CARGO_INCREMENTAL"
    printf '%s\n' 'allocator=CountingAllocator (process-local GlobalAlloc observer)'
    printf '%s\n' 'api_binding=reviewed XLSX source-backed owner adapter (set by wiring change)'
    printf '%s\n' "os=$(uname -srm 2>/dev/null || printf unavailable)"
    printf '%s\n' "cpu_model=$(awk -F: '/model name|Hardware/ {gsub(/^ +/, "", $2); print $2; exit}' /proc/cpuinfo 2>/dev/null || printf unavailable)"
    printf '%s\n' "core_count=$(nproc 2>/dev/null || printf unavailable)"
    printf '%s\n' "memory_total=$(awk '/MemTotal:/ {print $2 " " $3; exit}' /proc/meminfo 2>/dev/null || printf unavailable)"
    printf '%s\n' "storage=$(df -P "$ROOT" 2>/dev/null | tail -n 1 || printf unavailable)"
    printf '%s\n' "environment=CARGO_TARGET_DIR=$CARGO_TARGET_DIR CARGO_INCREMENTAL=$CARGO_INCREMENTAL"
} >"$RESULTS/build-provenance.txt"

: >"$RESULTS/commands.txt"
for lane in "${LANES[@]}"; do
    process=1
    while [[ "$process" -le "$PROCESSES" ]]; do
        report="$RESULTS/${lane}-p${process}.json"
        timing="$RESULTS/${lane}-p${process}.time.txt"
        printf '%s\n' \
            "run=/usr/bin/time -v $BIN --lane $lane --warmup $WARMUP --samples $SAMPLES (fresh_process=$process)" \
            >>"$RESULTS/commands.txt"
        /usr/bin/time -v -o "$timing" "$BIN" \
            --lane "$lane" --warmup "$WARMUP" --samples "$SAMPLES" \
            >"$report" 2>"$RESULTS/${lane}-p${process}.stderr.log"
        process=$((process + 1))
    done
done
sha256sum "$BIN" >"$RESULTS/binary-after.sha256"
cmp -s "$RESULTS/binary.sha256" "$RESULTS/binary-after.sha256"

cargo metadata --format-version=1 --locked --offline --manifest-path "$HARNESS" \
    >"$RESULTS/metadata-after.json"
python3 "$MANIFEST_TOOL" \
    --metadata "$RESULTS/metadata-after.json" \
    --root "$ROOT" \
    --output "$RESULTS/source-manifest-after.txt" \
    --extra "$ROOT/Cargo.toml" \
    --extra "$ROOT/rust-toolchain.toml" \
    --extra "$ROOT/.cargo/config.toml" \
    --extra "$HARNESS" \
    --extra "$HERE/harness/Cargo.lock" \
    --extra "$HERE/harness/adapter.rs" \
    --extra "$HERE/harness/support.rs" \
    --extra "$MANIFEST_TOOL" \
    --extra "$PROFILE_PINS" \
    --extra "$COMMITTED_INPUTS" \
    --extra "$HERE/run_profile.sh" \
    --extra "$HERE/summarize.py" \
    --extra "$HERE/uncertainty.py" \
    --extra "$HERE/verify.py" \
    --extra "$HERE/README.md" \
    --extra "$HERE/requirements.md" \
    --extra "$HERE/root-review.md" \
    --extra "$HERE/corpus-manifest.json" \
    --extra "$HERE/fixtures/tdf169496_hidden_graphic.xlsx" \
    --extra "$ROOT/docs/adr/0001-priorities-and-api-layers.md" \
    --extra "$ROOT/docs/adr/0005-io-memory-and-performance.md" \
    --extra "$ROOT/docs/report/spec-gap-validation-evidence/xlsx-svg-lifecycle-design.md" \
    --git-commit "$CURRENT_COMMIT"
cmp -s "$RESULTS/source-manifest-before.txt" "$RESULTS/source-manifest-after.txt"

{
    printf 'source_manifest_before_sha256='
    sha256sum "$RESULTS/source-manifest-before.txt" | cut -d' ' -f1
    printf 'source_manifest_after_sha256='
    sha256sum "$RESULTS/source-manifest-after.txt" | cut -d' ' -f1
    printf 'git_head='
    printf '%s\n' "$CURRENT_COMMIT"
    printf 'approved_source_pin=%s\n' "$SOURCE_PIN"
    printf '%s\n' 'committed_input_guard=profile_pins.py + committed_inputs.py + source_manifest.py'
    printf '%s\n' 'source_sha256:'
    for path in \
        "$ROOT/crates/litchi-xlsx/src/drawing/mod.rs" \
        "$ROOT/crates/litchi-xlsx/src/drawing/model.rs" \
        "$ROOT/crates/litchi-xlsx/src/drawing/codec.rs" \
        "$ROOT/crates/litchi-xlsx/src/workbook/edit/model.rs" \
        "$ROOT/crates/litchi-xlsx/src/workbook/edit/semantic/transaction.rs" \
        "$ROOT/crates/litchi-drawingml/src/svg_blip.rs"; do
        if [[ -f "$path" ]]; then sha256sum "$path"; fi
    done
    printf '%s\n' 'harness_sha256:'
    find "$HERE/harness" -type f -print0 | sort -z | xargs -0 sha256sum
    printf '%s\n' 'guard_sha256:'
    sha256sum "$PROFILE_PINS" "$COMMITTED_INPUTS" "$MANIFEST_TOOL"
} >"$RESULTS/source-provenance.txt"

python3 "$HERE/summarize.py" --results "$RESULTS" --output "$RESULTS/report.md"
python3 "$HERE/verify.py" --root "$ROOT" --results "$RESULTS" \
    --report "$RESULTS/report.md" --output "$RESULTS/verification.json"
python3 "$HERE/uncertainty.py" --results "$RESULTS" \
    --output "$RESULTS/uncertainty.md" --json-output "$RESULTS/uncertainty.json"
