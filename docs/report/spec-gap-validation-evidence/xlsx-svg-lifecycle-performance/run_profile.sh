#!/usr/bin/env bash
set -euo pipefail

if [[ "${PROFILE_FROZEN:-}" != "1" ]]; then
    echo "refusing to profile an unfrozen XLSX SVG lifecycle owner; set PROFILE_FROZEN=1" >&2
    exit 2
fi
if [[ "${XLSX_SVG_PROFILE_API_WIRED:-}" != "1" ]]; then
    echo "refusing to profile the unwired XLSX scaffold; set XLSX_SVG_PROFILE_API_WIRED=1 only after a reviewed API adapter is wired" >&2
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

HERE=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
ROOT=$(cd -- "$HERE/../../../.." && pwd)
HARNESS="$HERE/harness/Cargo.toml"
RESULTS="$HERE/results"
TARGET_INPUT=${CARGO_TARGET_DIR:-/var/tmp/litchi-xlsx-svg-lifecycle-profile-target}
if [[ "$TARGET_INPUT" = /* ]]; then
    TARGET=$(realpath -m -- "$TARGET_INPUT")
else
    TARGET=$(realpath -m -- "$PWD/$TARGET_INPUT")
fi
MANIFEST_TOOL="$HERE/source_manifest.py"
LANES=(
    capture_raster_two_cell_small capture_raster_two_cell_large
    capture_raster_one_cell_small capture_raster_one_cell_large
    capture_raster_absolute_small capture_raster_absolute_large
    capture_attached_two_cell_small capture_attached_two_cell_large
    capture_attached_one_cell_small capture_attached_one_cell_large
    capture_attached_absolute_small capture_attached_absolute_large
    clone_raster_small clone_raster_large
    clone_attached_small clone_attached_large
    inventory_shared_256 inventory_shared_1024
    inventory_distinct_256 inventory_distinct_1024
    namespace_heavy namespace_limit_refusal
    attach_end_to_end_two_cell_small attach_end_to_end_two_cell_large
    attach_end_to_end_one_cell_small attach_end_to_end_one_cell_large
    attach_end_to_end_absolute_small attach_end_to_end_absolute_large
    detach_end_to_end_shared_first_two_cell
    detach_end_to_end_shared_first_one_cell
    detach_end_to_end_shared_first_absolute
    detach_end_to_end_shared_final_two_cell
    detach_end_to_end_shared_final_one_cell
    detach_end_to_end_shared_final_absolute
    detach_end_to_end_distinct_two_cell_small
    detach_end_to_end_distinct_two_cell_large
    detach_end_to_end_distinct_one_cell_small
    detach_end_to_end_distinct_one_cell_large
    detach_end_to_end_distinct_absolute_small
    detach_end_to_end_distinct_absolute_large
    noop_detach_two_cell noop_detach_one_cell noop_detach_absolute
    limit_small limit_large
    malformed_duplicate_owner malformed_mce_owner
    malformed_linked_owner malformed_unknown_uri
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
if [[ "$TARGET" == "/" || "$TARGET" == "$ROOT" || "$TARGET" == "$ROOT/target" || "$TARGET" == "$HERE" ]]; then
    echo "refusing unsafe Cargo target path: $TARGET" >&2
    exit 2
fi

TARGET_CREATED=0
if [[ -e "$TARGET" && "${ALLOW_EXISTING_TARGET:-}" != "1" ]]; then
    echo "refusing to reuse an existing target; remove it or set ALLOW_EXISTING_TARGET=1" >&2
    exit 2
fi
if [[ ! -e "$TARGET" ]]; then
    TARGET_CREATED=1
fi
cleanup_target() {
    if [[ "$TARGET_CREATED" == "1" ]]; then
        rm -rf -- "$TARGET"
    fi
}
trap cleanup_target EXIT

mkdir -p "$RESULTS" "$TARGET"
rm -f "$RESULTS"/*.json "$RESULTS"/*.time.txt "$RESULTS"/*.log \
    "$RESULTS"/commands.txt "$RESULTS"/source-manifest-*.txt \
    "$RESULTS"/source-provenance.txt "$RESULTS"/binary.sha256 \
    "$RESULTS"/build-provenance.txt "$HERE/report.md" "$HERE/verification.json"

unset RUSTFLAGS CARGO_ENCODED_RUSTFLAGS RUSTC_BOOTSTRAP
export CARGO_TARGET_DIR="$TARGET"
export CARGO_INCREMENTAL=0

cargo metadata --format-version=1 --locked --offline --manifest-path "$HARNESS" \
    >"$RESULTS/metadata-before.json"
python3 "$MANIFEST_TOOL" \
    --metadata "$RESULTS/metadata-before.json" \
    --root "$ROOT" \
    --output "$RESULTS/source-manifest-before.txt" \
    --extra "$ROOT/Cargo.toml" \
    --extra "$ROOT/Cargo.lock" \
    --extra "$HARNESS" \
    --extra "$HERE/harness/Cargo.lock" \
    --extra "$HERE/harness/support.rs" \
    --extra "$MANIFEST_TOOL" \
    --extra "$HERE/run_profile.sh" \
    --extra "$HERE/summarize.py" \
    --extra "$HERE/verify.py" \
    --extra "$HERE/README.md" \
    --extra "$HERE/requirements.md" \
    --extra "$HERE/root-review.md" \
    --extra "$HERE/corpus-manifest.json" \
    --extra "$ROOT/docs/GOAL.md" \
    --extra "$ROOT/docs/report/spec-gap-validation-evidence/xlsx-svg-lifecycle-design.md"

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
    printf '%s\n' "cpu_model=$(awk -F: '/model name|Hardware/ {gsub(/^ +/, \"\", $2); print $2; exit}' /proc/cpuinfo 2>/dev/null || printf unavailable)"
    printf '%s\n' "core_count=$(nproc 2>/dev/null || printf unavailable)"
    printf '%s\n' "memory_total=$(awk '/MemTotal:/ {print $2 \" \" $3; exit}' /proc/meminfo 2>/dev/null || printf unavailable)"
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
            >"$report"
        process=$((process + 1))
    done
done

cargo metadata --format-version=1 --locked --offline --manifest-path "$HARNESS" \
    >"$RESULTS/metadata-after.json"
python3 "$MANIFEST_TOOL" \
    --metadata "$RESULTS/metadata-after.json" \
    --root "$ROOT" \
    --output "$RESULTS/source-manifest-after.txt" \
    --extra "$ROOT/Cargo.toml" \
    --extra "$ROOT/Cargo.lock" \
    --extra "$HARNESS" \
    --extra "$HERE/harness/Cargo.lock" \
    --extra "$HERE/harness/support.rs" \
    --extra "$MANIFEST_TOOL" \
    --extra "$HERE/run_profile.sh" \
    --extra "$HERE/summarize.py" \
    --extra "$HERE/verify.py" \
    --extra "$HERE/README.md" \
    --extra "$HERE/requirements.md" \
    --extra "$HERE/root-review.md" \
    --extra "$HERE/corpus-manifest.json" \
    --extra "$ROOT/docs/GOAL.md" \
    --extra "$ROOT/docs/report/spec-gap-validation-evidence/xlsx-svg-lifecycle-design.md"
cmp -s "$RESULTS/source-manifest-before.txt" "$RESULTS/source-manifest-after.txt"

{
    printf 'source_manifest_before_sha256='
    sha256sum "$RESULTS/source-manifest-before.txt" | cut -d' ' -f1
    printf 'source_manifest_after_sha256='
    sha256sum "$RESULTS/source-manifest-after.txt" | cut -d' ' -f1
    printf 'git_head='
    git -C "$ROOT" rev-parse HEAD
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
} >"$RESULTS/source-provenance.txt"

python3 "$HERE/summarize.py" --results "$RESULTS" --output "$HERE/report.md"
python3 "$HERE/verify.py"
