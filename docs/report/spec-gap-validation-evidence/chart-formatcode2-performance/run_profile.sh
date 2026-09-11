#!/bin/sh

set -eu

if [ "${PROFILE_FROZEN:-0}" != 1 ]; then
    echo "profile is intentionally gated: set PROFILE_FROZEN=1 only after the formatcode2 owner and semantic tests are frozen" >&2
    exit 2
fi

HERE=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
ROOT=$(CDPATH= cd -- "$HERE/../../../../" && pwd)
HARNESS="$HERE/harness/Cargo.toml"
MANIFEST_TOOL="$HERE/source_manifest.py"
RESULTS="$HERE/results"
TARGET=${CARGO_TARGET_DIR:-/var/tmp/litchi-chart-formatcode2-profile-target}
WARMUP=${WARMUP:-2}
SAMPLES=${SAMPLES:-20}
PROCESSES=${PROCESSES:-3}
LANES="element_small_read element_small_noop element_small_scalar_edit element_small_clone element_small_read_shared element_small_write_to element_small_malformed element_near_limit_read element_near_limit_noop element_near_limit_scalar_edit element_near_limit_clone element_near_limit_read_shared element_near_limit_write_to element_near_limit_malformed attribute_small_read attribute_small_noop attribute_small_scalar_edit attribute_small_clone attribute_small_read_shared attribute_small_write_to attribute_small_malformed attribute_near_limit_read attribute_near_limit_noop attribute_near_limit_scalar_edit attribute_near_limit_clone attribute_near_limit_read_shared attribute_near_limit_write_to attribute_near_limit_malformed"

for count in "$WARMUP" "$SAMPLES" "$PROCESSES"; do
    case "$count" in
        ''|*[!0-9]*|0) echo "WARMUP, SAMPLES, and PROCESSES must be positive integers" >&2; exit 2 ;;
    esac
done

for path in \
    "$ROOT/crates/litchi-drawingml/src/chart/extension/formatcode2.rs" \
    "$ROOT/crates/litchi-drawingml/tests/chart_formatcode2.rs" \
    "$ROOT/crates/litchi-drawingml/tests/chart_formatcode2_adversarial.rs"; do
    if [ ! -f "$path" ]; then
        echo "formatcode2 source input is missing: $path" >&2
        exit 1
    fi
done

mkdir -p "$RESULTS"
TARGET_OWNED=0
if [ ! -e "$TARGET" ]; then
    TARGET_OWNED=1
fi
TEMP_DIR=$(mktemp -d "${TMPDIR:-/tmp}/litchi-chart-formatcode2-profile.XXXXXX")

cleanup() {
    status=$?
    rm -rf -- "$TEMP_DIR"
    if [ "${CLEAN_TARGET:-1}" = 1 ] && [ "$TARGET_OWNED" -eq 1 ]; then
        case "$TARGET" in
            /var/tmp/litchi-chart-formatcode2-profile-target|/tmp/litchi-chart-formatcode2-profile-target)
                rm -rf -- "$TARGET" ;;
        esac
    fi
    exit "$status"
}
trap cleanup EXIT

rm -f "$RESULTS"/*.json "$RESULTS"/*.time.txt "$RESULTS"/source-* \
    "$RESULTS"/binary.sha256 "$RESULTS"/build.log "$RESULTS"/build-provenance.txt \
    "$RESULTS"/commands.txt "$RESULTS"/host.txt "$RESULTS"/verification.json

{
    printf '%s\n' "root=$ROOT"
    printf '%s\n' "harness=$HARNESS"
    printf '%s\n' "target=$TARGET"
    printf '%s\n' "warmup=$WARMUP"
    printf '%s\n' "samples_per_process=$SAMPLES"
    printf '%s\n' "fresh_processes_per_lane=$PROCESSES"
    printf '%s\n' "lanes=$LANES"
    printf '%s\n' "command=cargo metadata --format-version=1 --locked --manifest-path $HARNESS"
    printf '%s\n' "command=CARGO_INCREMENTAL=0 cargo build --release --locked --manifest-path $HARNESS"
    printf '%s\n' "command=/usr/bin/time -v $TARGET/release/chart-formatcode2-profile --lane {lane} --warmup $WARMUP --samples $SAMPLES"
} >"$RESULTS/commands.txt"

{
    date -u '+utc=%Y-%m-%dT%H:%M:%SZ'
    rustc -vV
    cargo -V
    uname -a
    if command -v lscpu >/dev/null 2>&1; then lscpu | sed -n '1,32p'; fi
} >"$RESULTS/host.txt"

export CARGO_TARGET_DIR="$TARGET"
export CARGO_INCREMENTAL=0

cargo metadata --format-version=1 --locked --manifest-path "$HARNESS" >"$TEMP_DIR/metadata-before.json"
python3 "$MANIFEST_TOOL" \
    --metadata "$TEMP_DIR/metadata-before.json" \
    --root "$ROOT" \
    --output "$RESULTS/source-manifest-before.txt" \
    --extra "$ROOT/Cargo.toml" \
    --extra "$ROOT/Cargo.lock" \
    --extra "$HARNESS" \
    --extra "$HERE/harness/Cargo.lock" \
    --extra "$MANIFEST_TOOL" \
    --extra "$HERE/run_profile.sh" \
    --extra "$HERE/summarize.py" \
    --extra "$HERE/verify.py"

cargo build --release --locked --manifest-path "$HARNESS" >"$RESULTS/build.log" 2>&1
BIN="$TARGET/release/chart-formatcode2-profile"
if [ ! -x "$BIN" ]; then echo "profile binary was not produced: $BIN" >&2; exit 1; fi
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

for lane in $LANES; do
    process=1
    while [ "$process" -le "$PROCESSES" ]; do
        report="$RESULTS/${lane}-p${process}.json"
        timing="$RESULTS/${lane}-p${process}.time.txt"
        printf '%s\n' "run=/usr/bin/time -v $BIN --lane $lane --warmup $WARMUP --samples $SAMPLES (fresh_process=$process)" >>"$RESULTS/commands.txt"
        /usr/bin/time -v -o "$timing" "$BIN" \
            --lane "$lane" \
            --warmup "$WARMUP" \
            --samples "$SAMPLES" >"$report"
        process=$((process + 1))
    done
done

cargo metadata --format-version=1 --locked --manifest-path "$HARNESS" >"$TEMP_DIR/metadata-after.json"
python3 "$MANIFEST_TOOL" \
    --metadata "$TEMP_DIR/metadata-after.json" \
    --root "$ROOT" \
    --output "$RESULTS/source-manifest-after.txt" \
    --extra "$ROOT/Cargo.toml" \
    --extra "$ROOT/Cargo.lock" \
    --extra "$HARNESS" \
    --extra "$HERE/harness/Cargo.lock" \
    --extra "$MANIFEST_TOOL" \
    --extra "$HERE/run_profile.sh" \
    --extra "$HERE/summarize.py" \
    --extra "$HERE/verify.py"
cmp -s "$RESULTS/source-manifest-before.txt" "$RESULTS/source-manifest-after.txt"

{
    printf 'source_manifest_before_sha256='
    sha256sum "$RESULTS/source-manifest-before.txt" | cut -d' ' -f1
    printf 'source_manifest_after_sha256='
    sha256sum "$RESULTS/source-manifest-after.txt" | cut -d' ' -f1
    printf 'git_head='
    git -C "$ROOT" rev-parse HEAD
    printf 'git_status_relevant='
    git -C "$ROOT" status --short -- \
        crates/litchi-drawingml/src/chart crates/litchi-drawingml/tests/chart_formatcode2.rs \
        crates/litchi-drawingml/tests/chart_formatcode2_adversarial.rs | tr '\n' '|'
    printf '\nsource_sha256:\n'
    for path in \
        "$ROOT/crates/litchi-drawingml/src/chart/mod.rs" \
        "$ROOT/crates/litchi-drawingml/src/chart/extension/mod.rs" \
        "$ROOT/crates/litchi-drawingml/src/chart/extension/formatcode2.rs" \
        "$ROOT/crates/litchi-drawingml/tests/chart_formatcode2.rs" \
        "$ROOT/crates/litchi-drawingml/tests/chart_formatcode2_adversarial.rs"; do
        sha256sum "$path"
    done
    printf '%s\n' 'harness_sha256:'
    find "$HERE/harness" -type f -print0 | sort -z | xargs -0 sha256sum
} >"$RESULTS/source-provenance.txt"

python3 "$HERE/summarize.py" --results "$RESULTS" --output "$HERE/report.md"
python3 "$HERE/verify.py"
