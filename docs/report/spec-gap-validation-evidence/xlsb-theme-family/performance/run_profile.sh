#!/bin/sh

set -eu

HERE=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
ROOT=$(CDPATH= cd -- "$HERE/../../../../.." && pwd)
HARNESS="$HERE/harness/Cargo.toml"
MANIFEST_TOOL="$HERE/source_manifest.py"
TARGET=${CARGO_TARGET_DIR:-/var/tmp/litchi-xlsb-theme-family-profile-target}
RESULTS=${RESULTS_DIR:-$HERE/results}
FIXTURE="$ROOT/test-data/ooxml/xlsb/date.xlsb"
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

if [ ! -f "$FIXTURE" ]; then
    echo "native XLSB fixture is missing: $FIXTURE" >&2
    exit 1
fi

mkdir -p "$RESULTS"
PROFILE_LOCK="$RESULTS/.profile.lock"
if ! mkdir "$PROFILE_LOCK" 2>/dev/null; then
    echo "another XLSB theme-family profile is writing $RESULTS" >&2
    exit 1
fi

TARGET_OWNED=0
if [ ! -e "$TARGET" ]; then
    TARGET_OWNED=1
fi
TEMP_DIR=$(mktemp -d "${TMPDIR:-/tmp}/litchi-xlsb-theme-family-profile.XXXXXX")

cleanup() {
    status=$?
    rm -rf -- "$TEMP_DIR"
    rmdir "$PROFILE_LOCK" 2>/dev/null || true
    if [ "${CLEAN_TARGET:-1}" = 1 ] && [ "$TARGET_OWNED" -eq 1 ]; then
        case "$TARGET" in
            /var/tmp/litchi-xlsb-theme-family-profile-target|/tmp/litchi-xlsb-theme-family-profile-target)
                rm -rf -- "$TARGET"
                ;;
        esac
    fi
    exit "$status"
}
trap cleanup EXIT

for fixture_kind in native opaque; do
    for operation in codec_read metadata_read source_read family_clone noop add update remove base_edit; do
        rm -f \
            "$RESULTS/${fixture_kind}-${operation}-p"*.json \
            "$RESULTS/${fixture_kind}-${operation}-p"*.time.txt
    done
done
rm -f \
    "$RESULTS/perf-opaque-update.csv" \
    "$RESULTS/perf-opaque-update.json" \
    "$RESULTS/perf-opaque-update.stderr" \
    "$RESULTS/source-manifest-before.txt" \
    "$RESULTS/source-manifest-after.txt" \
    "$RESULTS/source-manifest.txt" \
    "$RESULTS/source-manifest.diff" \
    "$RESULTS/source-provenance.txt" \
    "$RESULTS/build-provenance.txt" \
    "$RESULTS/binary.sha256" \
    "$RESULTS/commands.txt" \
    "$RESULTS/host.txt" \
    "$RESULTS/build.log"

{
    printf '%s\n' "root=$ROOT"
    printf '%s\n' "harness=$HARNESS"
    printf '%s\n' "fixture=$FIXTURE"
    printf '%s\n' "target=$TARGET"
    printf '%s\n' "warmup=$WARMUP"
    printf '%s\n' "samples_per_process=$SAMPLES"
    printf '%s\n' "fresh_processes_per_lane=$PROCESSES"
    printf '%s\n' "allocator_instrumented=true"
    printf '%s\n' "timed_input_clone_excluded=true for metadata_read; source/pointer/hash checks are outside timed closures"
    printf '%s\n' "command=cargo metadata --format-version=1 --locked --manifest-path $HARNESS"
    printf '%s\n' "command=python3 $MANIFEST_TOOL --metadata metadata --root $ROOT --output $RESULTS/source-manifest-before.txt"
    printf '%s\n' "command=CARGO_INCREMENTAL=0 cargo build --release --locked --manifest-path $HARNESS"
    printf '%s\n' "command=/usr/bin/time -v $TARGET/release/xlsb-theme-family-profile --fixture $FIXTURE --kind {native,opaque} --operation {codec_read,metadata_read,source_read,family_clone,noop,add,update,remove,base_edit} --warmup $WARMUP --samples $SAMPLES"
    printf '%s\n' "command=python3 $MANIFEST_TOOL --metadata metadata --root $ROOT --output $RESULTS/source-manifest-after.txt"
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
    --extra "$HERE/summarize.py" \
    --extra "$HERE/verify_root.py" \
    --extra "$FIXTURE"

cargo build --release --locked --manifest-path "$HARNESS" >"$RESULTS/build.log" 2>&1
BIN="$TARGET/release/xlsb-theme-family-profile"
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
    printf '%s\n' "allocator=CountingAllocator (process-local GlobalAlloc observer)"
} >"$RESULTS/build-provenance.txt"

for fixture_kind in native opaque; do
    for operation in codec_read metadata_read source_read family_clone noop add update remove base_edit; do
        process=1
        while [ "$process" -le "$PROCESSES" ]; do
            report="$RESULTS/${fixture_kind}-${operation}-p${process}.json"
            timing="$RESULTS/${fixture_kind}-${operation}-p${process}.time.txt"
            printf '%s\n' "run=/usr/bin/time -v $BIN --fixture $FIXTURE --kind $fixture_kind --operation $operation --warmup $WARMUP --samples $SAMPLES (fresh_process=$process)" >>"$RESULTS/commands.txt"
            /usr/bin/time -v -o "$timing" "$BIN" \
                --fixture "$FIXTURE" \
                --kind "$fixture_kind" \
                --operation "$operation" \
                --warmup "$WARMUP" \
                --samples "$SAMPLES" >"$report"
            process=$((process + 1))
        done
    done
done

if [ "${PERF:-0}" = 1 ]; then
    printf '%s\n' "run=perf stat -x, -e cycles,instructions,branches,branch-misses,cache-misses $BIN --fixture $FIXTURE --kind opaque --operation update --warmup $WARMUP --samples $SAMPLES" >>"$RESULTS/commands.txt"
    set +e
    perf stat -x, -e cycles,instructions,branches,branch-misses,cache-misses \
        -o "$RESULTS/perf-opaque-update.csv" \
        "$BIN" --fixture "$FIXTURE" --kind opaque --operation update \
        --warmup "$WARMUP" --samples "$SAMPLES" \
        >"$RESULTS/perf-opaque-update.json" 2>"$RESULTS/perf-opaque-update.stderr"
    status=$?
    set -e
    printf 'perf_exit=%s\n' "$status" >>"$RESULTS/commands.txt"
fi

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
    --extra "$HERE/summarize.py" \
    --extra "$HERE/verify_root.py" \
    --extra "$FIXTURE"
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
        crates/litchi-drawingml/src/theme \
        crates/litchi-xlsb/src/theme.rs \
        crates/litchi-xlsb/src/theme/lifecycle.rs \
        crates/litchi-xlsb/src/writer/workbook/model.rs \
        crates/litchi-xlsb/src/writer/workbook/package.rs \
        "$HERE" | tr '\n' '|'; printf '\n'
    printf '%s\n' 'fixture_sha256:'
    sha256sum "$FIXTURE"
    printf '%s\n' 'dirty_source_sha256:'
    for path in \
        "$ROOT/crates/litchi-drawingml/src/theme/mod.rs" \
        "$ROOT/crates/litchi-xlsb/src/theme.rs" \
        "$ROOT/crates/litchi-xlsb/src/theme/lifecycle.rs" \
        "$ROOT/crates/litchi-xlsb/src/writer/workbook/model.rs" \
        "$ROOT/crates/litchi-xlsb/src/writer/workbook/package.rs"; do
        sha256sum "$path"
    done
    find "$ROOT/crates/litchi-drawingml/src/theme/family" -type f -print | sort | while IFS= read -r path; do
        sha256sum "$path"
    done
    printf '%s\n' 'validation_source_sha256:'
    sha256sum \
        "$ROOT/crates/litchi-drawingml/tests/theme_part.rs" \
        "$ROOT/crates/litchi-drawingml/tests/fixtures/theme-part-native.xml" \
        "$ROOT/crates/litchi-drawingml/tests/fixtures/theme-part-native.provenance" \
        "$ROOT/crates/litchi-xlsb/tests/theme_family.rs"
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
if [ "$PROCESSES" = 3 ] && [ "$SAMPLES" = 30 ]; then
    python3 "$HERE/verify_root.py"
fi
