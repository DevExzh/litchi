#!/bin/sh

set -eu

HERE=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
ROOT=$(CDPATH= cd -- "$HERE/../../../.." && pwd)
HARNESS="$HERE/harness/Cargo.toml"
MANIFEST_TOOL="$HERE/source_manifest.py"
TARGET=${CARGO_TARGET_DIR:-/var/tmp/litchi-drawingml-theme-family-profile-target}
RESULTS=${RESULTS_DIR:-$HERE/results}
WARMUP=${WARMUP:-3}
SAMPLES=${SAMPLES:-30}
PROCESSES=${PROCESSES:-3}

for count in "$WARMUP" "$SAMPLES" "$PROCESSES"; do
    case "$count" in
        ''|*[!0-9]*|0)
            echo "WARMUP, SAMPLES, and PROCESSES must be positive integers" >&2
            exit 2
            ;;
    esac
done

mkdir -p "$RESULTS"
PROFILE_LOCK="$RESULTS/.profile.lock"
if ! mkdir "$PROFILE_LOCK" 2>/dev/null; then
    echo "another theme-family profile is writing $RESULTS" >&2
    exit 1
fi

# A caller-owned target directory is reusable and must survive this run. The
# default target is removed only when this invocation created it.
TARGET_OWNED=0
if [ ! -e "$TARGET" ]; then
    TARGET_OWNED=1
fi

TEMP_DIR=$(mktemp -d "${TMPDIR:-/tmp}/drawingml-theme-family-profile.XXXXXX")

cleanup() {
    status=$?
    rm -rf -- "$TEMP_DIR"
    rmdir "$PROFILE_LOCK" 2>/dev/null || true
    if [ "${CLEAN_TARGET:-1}" = 1 ] && [ "$TARGET_OWNED" -eq 1 ]; then
        case "$TARGET" in
            /var/tmp/litchi-drawingml-theme-family-profile-target|/tmp/litchi-drawingml-theme-family-profile-target)
                rm -rf -- "$TARGET"
                ;;
        esac
    fi
    exit "$status"
}
trap cleanup EXIT

for fixture in small opaque; do
    for operation in read clone noop change; do
        rm -f \
            "$RESULTS/${fixture}-${operation}.json" \
            "$RESULTS/${fixture}-${operation}.time.txt"
        find "$RESULTS" -maxdepth 1 -type f \
            -name "${fixture}-${operation}-p*.json" -delete
        find "$RESULTS" -maxdepth 1 -type f \
            -name "${fixture}-${operation}-p*.time.txt" -delete
    done
done
rm -f \
    "$RESULTS/perf-opaque-change.csv" \
    "$RESULTS/perf-opaque-change.json" \
    "$RESULTS/perf-opaque-change.stderr" \
    "$RESULTS/source-manifest-before.txt" \
    "$RESULTS/source-manifest-after.txt" \
    "$RESULTS/source-manifest.txt" \
    "$RESULTS/source-manifest.diff" \
    "$RESULTS/source-provenance.txt" \
    "$RESULTS/build-provenance.txt" \
    "$RESULTS/binary.sha256"

{
    printf '%s\n' "root=$ROOT"
    printf '%s\n' "harness=$HARNESS"
    printf '%s\n' "target=$TARGET"
    printf '%s\n' "warmup=$WARMUP"
    printf '%s\n' "samples_per_process=$SAMPLES"
    printf '%s\n' "fresh_processes_per_lane=$PROCESSES"
    printf '%s\n' "allocator_instrumented=true"
    printf '%s\n' "command=cargo metadata --format-version=1 --locked --manifest-path $HARNESS > $TEMP_DIR/metadata-before.json"
    printf '%s\n' "command=python3 $MANIFEST_TOOL --metadata $TEMP_DIR/metadata-before.json --root $ROOT --output $RESULTS/source-manifest-before.txt --extra $ROOT/Cargo.toml $ROOT/Cargo.lock $HARNESS $HERE/harness/Cargo.lock $MANIFEST_TOOL $HERE/run_profile.sh $HERE/summarize.py"
    printf '%s\n' "command=CARGO_INCREMENTAL=0 cargo build --release --locked --manifest-path $HARNESS"
    printf '%s\n' "command=cargo metadata --format-version=1 --locked --manifest-path $HARNESS > $TEMP_DIR/metadata-after.json"
    printf '%s\n' "command=python3 $MANIFEST_TOOL --metadata $TEMP_DIR/metadata-after.json --root $ROOT --output $RESULTS/source-manifest-after.txt --extra $ROOT/Cargo.toml $ROOT/Cargo.lock $HARNESS $HERE/harness/Cargo.lock $MANIFEST_TOOL $HERE/run_profile.sh $HERE/summarize.py"
} >"$RESULTS/commands.txt"

{
    date -u '+utc=%Y-%m-%dT%H:%M:%SZ'
    rustc -vV
    cargo -V
    uname -a
    if command -v lscpu >/dev/null 2>&1; then
        lscpu | sed -n '1,32p'
    fi
} >"$RESULTS/host.txt"

export CARGO_TARGET_DIR="$TARGET"
export CARGO_INCREMENTAL=0

cargo metadata --format-version=1 --locked --manifest-path "$HARNESS" \
    >"$TEMP_DIR/metadata-before.json"
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
    --extra "$HERE/summarize.py"

cargo build --release --locked --manifest-path "$HARNESS" >"$RESULTS/build.log" 2>&1
BIN="$TARGET/release/drawingml-theme-family-profile"
if [ ! -x "$BIN" ]; then
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
} >"$RESULTS/build-provenance.txt"

for fixture in small opaque; do
    for operation in read clone noop change; do
        process=1
        while [ "$process" -le "$PROCESSES" ]; do
            report="$RESULTS/${fixture}-${operation}-p${process}.json"
            timing="$RESULTS/${fixture}-${operation}-p${process}.time.txt"
            printf '%s\n' "run=/usr/bin/time -v $BIN --fixture $fixture --operation $operation --warmup $WARMUP --samples $SAMPLES (fresh_process=$process)" >>"$RESULTS/commands.txt"
            /usr/bin/time -v -o "$timing" "$BIN" \
                --fixture "$fixture" \
                --operation "$operation" \
                --warmup "$WARMUP" \
                --samples "$SAMPLES" >"$report"
            process=$((process + 1))
        done
    done
done

if [ "${PERF:-0}" = 1 ]; then
    printf '%s\n' "run=perf stat -x, -e cycles,instructions,branches,branch-misses,cache-misses $BIN --fixture opaque --operation change --warmup $WARMUP --samples $SAMPLES" >>"$RESULTS/commands.txt"
    set +e
    perf stat -x, -e cycles,instructions,branches,branch-misses,cache-misses \
        -o "$RESULTS/perf-opaque-change.csv" \
        "$BIN" --fixture opaque --operation change --warmup "$WARMUP" --samples "$SAMPLES" \
        >"$RESULTS/perf-opaque-change.json" 2>"$RESULTS/perf-opaque-change.stderr"
    status=$?
    set -e
    printf 'perf_exit=%s\n' "$status" >>"$RESULTS/commands.txt"
fi

# Re-read and re-hash the complete Cargo package source set after both build
# and execution. A mutation would make the binary/profiles non-reproducible.
cargo metadata --format-version=1 --locked --manifest-path "$HARNESS" \
    >"$TEMP_DIR/metadata-after.json"
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
    --extra "$HERE/summarize.py"
if ! cmp -s "$RESULTS/source-manifest-before.txt" "$RESULTS/source-manifest-after.txt"; then
    diff -u "$RESULTS/source-manifest-before.txt" "$RESULTS/source-manifest-after.txt" \
        >"$RESULTS/source-manifest.diff" || true
    echo "Cargo source or manifest inputs changed during the profile" >&2
    exit 1
fi

before_hash=$(sha256sum "$RESULTS/source-manifest-before.txt" | cut -d' ' -f1)
after_hash=$(sha256sum "$RESULTS/source-manifest-after.txt" | cut -d' ' -f1)
{
    printf 'source_manifest_before_sha256=%s\n' "$before_hash"
    printf 'source_manifest_after_sha256=%s\n' "$after_hash"
    printf '%s\n' 'source_manifest_match=true'
    printf 'git_head='; git -C "$ROOT" rev-parse HEAD
    printf 'git_status_relevant='; git -C "$ROOT" status --short -- \
        crates/litchi-drawingml/src/theme/mod.rs \
        crates/litchi-drawingml/src/theme/family \
        crates/litchi-drawingml/tests/theme_family.rs \
        crates/litchi-drawingml/tests/fixtures/theme-family-native.xml \
        crates/litchi-drawingml/tests/fixtures/theme-family-native.provenance | tr '\n' '|'; printf '\n'
    printf '%s\n' 'dirty_theme_source_sha256:'
    sha256sum "$ROOT/crates/litchi-drawingml/src/theme/mod.rs"
    find "$ROOT/crates/litchi-drawingml/src/theme/family" -type f -print | sort | while IFS= read -r path; do
        sha256sum "$path"
    done
    printf '%s\n' 'validation_source_sha256:'
    sha256sum \
        "$ROOT/crates/litchi-drawingml/tests/theme_family.rs" \
        "$ROOT/crates/litchi-drawingml/tests/fixtures/theme-family-native.xml" \
        "$ROOT/crates/litchi-drawingml/tests/fixtures/theme-family-native.provenance"
} >"$RESULTS/source-provenance.txt"
{
    cat "$RESULTS/source-provenance.txt"
    printf '%s\n' '--- Cargo package source manifest after build/run ---'
    cat "$RESULTS/source-manifest-after.txt"
} >"$RESULTS/source-manifest.txt"

python3 "$HERE/summarize.py" \
    --results "$RESULTS" \
    --output "$HERE/report.md" \
    --expected-processes "$PROCESSES" \
    --samples-per-process "$SAMPLES"
