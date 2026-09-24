#!/usr/bin/env bash
set -euo pipefail

if [[ "${PROFILE_FROZEN:-}" != "1" ]]; then
    echo "refusing to profile an unfrozen SVG lifecycle owner; set PROFILE_FROZEN=1" >&2
    exit 2
fi
for variable in RUSTFLAGS CARGO_ENCODED_RUSTFLAGS RUSTC_BOOTSTRAP; do
    if [[ -n "${!variable:-}" ]]; then
        echo "refusing profile with ${variable} set" >&2
        exit 2
    fi
done

HERE=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
ROOT=$(cd -- "$HERE/../../../.." && pwd)
HARNESS="$HERE/harness/Cargo.toml"
RESULTS="$HERE/results"
TARGET=${CARGO_TARGET_DIR:-/var/tmp/litchi-pptx-svg-lifecycle-profile-target}
MANIFEST_TOOL="$HERE/source_manifest.py"
LANES=(
    capture_raster_small capture_raster_large
    capture_attached_small capture_attached_large
    capture_namespace_heavy
    inventory_many_raster_256 inventory_many_raster_1024
    inventory_distinct_local_namespace_256 inventory_distinct_local_namespace_1024
    attach_end_to_end_small attach_end_to_end_large
    detach_end_to_end_small detach_end_to_end_large
    noop_detach_end_to_end_small noop_detach_end_to_end_large
    clone_raster_small clone_raster_large
    clone_attached_small clone_attached_large
    limit_small limit_large
    malformed_small malformed_large
    namespace_limit_refusal
)
PROCESSES=${PROCESSES:-3}
WARMUP=${WARMUP:-2}
SAMPLES=${SAMPLES:-20}

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
    --extra "$MANIFEST_TOOL" \
    --extra "$HERE/run_profile.sh" \
    --extra "$HERE/summarize.py" \
    --extra "$HERE/verify.py" \
    --extra "$HERE/README.md" \
    --extra "$HERE/requirements.md"

cargo build --release --locked --offline --manifest-path "$HARNESS" \
    >"$RESULTS/build.log" 2>&1
BIN="$TARGET/release/pptx-svg-lifecycle-profile"
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
    --extra "$MANIFEST_TOOL" \
    --extra "$HERE/run_profile.sh" \
    --extra "$HERE/summarize.py" \
    --extra "$HERE/verify.py" \
    --extra "$HERE/README.md" \
    --extra "$HERE/requirements.md"
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
        "$ROOT/crates/litchi-pptx/src/presentation/source.rs" \
        "$ROOT/crates/litchi-pptx/src/presentation/source/svg_lifecycle.rs" \
        "$ROOT/crates/litchi-pptx/src/presentation/source/svg_lifecycle/owner.rs" \
        "$ROOT/crates/litchi-opc/src/source_backed.rs" \
        "$ROOT/crates/litchi-opc/src/phys_pkg.rs" \
        "$ROOT/crates/litchi-pptx/src/lib.rs" \
        "$ROOT/crates/litchi-pptx/src/presentation/mod.rs"; do
        if [[ -f "$path" ]]; then sha256sum "$path"; fi
    done
    printf '%s\n' 'harness_sha256:'
    find "$HERE/harness" -type f -print0 | sort -z | xargs -0 sha256sum
} >"$RESULTS/source-provenance.txt"

python3 "$HERE/summarize.py" --results "$RESULTS" --output "$HERE/report.md"
python3 "$HERE/verify.py"
