#!/bin/sh

set -eu

HERE=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
ROOT=$(CDPATH= cd -- "$HERE/../../../../" && pwd)
HARNESS="$HERE/harness/Cargo.toml"
MANIFEST_TOOL="$HERE/source_manifest.py"
FIXTURE="$ROOT/crates/litchi-drawingml/tests/fixtures/theme-part-native.xml"
TARGET=${CARGO_TARGET_DIR:-/var/tmp/litchi-theme-family-hardening-profile-target}
RESULTS="$HERE/results"
WARMUP=${WARMUP:-2}
SAMPLES=${SAMPLES:-20}
PROCESSES=${PROCESSES:-3}
LANES="native_read native_replace native_remove native_add unknown_32 unknown_1000 duplicate limit_replace limit_add"

for count in "$WARMUP" "$SAMPLES" "$PROCESSES"; do
    case "$count" in
        ''|*[!0-9]*|0) echo "WARMUP, SAMPLES, and PROCESSES must be positive integers" >&2; exit 2 ;;
    esac
done

if [ ! -f "$FIXTURE" ]; then
    echo "native Theme fixture is missing: $FIXTURE" >&2
    exit 1
fi

mkdir -p "$RESULTS"
TARGET_OWNED=0
if [ ! -e "$TARGET" ]; then
    TARGET_OWNED=1
fi
TEMP_DIR=$(mktemp -d "${TMPDIR:-/tmp}/litchi-theme-family-hardening.XXXXXX")

cleanup() {
    status=$?
    rm -rf -- "$TEMP_DIR"
    if [ "${CLEAN_TARGET:-1}" = 1 ] && [ "$TARGET_OWNED" -eq 1 ]; then
        case "$TARGET" in
            /var/tmp/litchi-theme-family-hardening-profile-target|/tmp/litchi-theme-family-hardening-profile-target)
                rm -rf -- "$TARGET" ;;
        esac
    fi
    exit "$status"
}
trap cleanup EXIT

rm -f "$RESULTS"/*.json "$RESULTS"/*.time.txt "$RESULTS"/source-* \
    "$RESULTS"/binary.sha256 "$RESULTS"/build.log "$RESULTS"/build-provenance.txt \
    "$RESULTS"/commands.txt "$RESULTS"/host.txt "$RESULTS"/source-provenance.txt

{
    printf '%s\n' "root=$ROOT"
    printf '%s\n' "harness=$HARNESS"
    printf '%s\n' "fixture=$FIXTURE"
    printf '%s\n' "target=$TARGET"
    printf '%s\n' "warmup=$WARMUP"
    printf '%s\n' "samples_per_process=$SAMPLES"
    printf '%s\n' "fresh_processes_per_lane=$PROCESSES"
    printf '%s\n' "lanes=$LANES"
    printf '%s\n' "command=cargo metadata --format-version=1 --locked --manifest-path $HARNESS"
    printf '%s\n' "command=CARGO_INCREMENTAL=0 cargo build --release --locked --manifest-path $HARNESS"
    printf '%s\n' "command=/usr/bin/time -v $TARGET/release/theme-family-hardening-profile --fixture $FIXTURE --lane {lane} --warmup $WARMUP --samples $SAMPLES"
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
    --extra "$HERE/run_targeted.sh" \
    --extra "$HERE/summarize_targeted.py" \
    --extra "$HERE/verify.py" \
    --extra "$FIXTURE"

cargo build --release --locked --manifest-path "$HARNESS" >"$RESULTS/build.log" 2>&1
BIN="$TARGET/release/theme-family-hardening-profile"
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
        printf '%s\n' "run=/usr/bin/time -v $BIN --fixture $FIXTURE --lane $lane --warmup $WARMUP --samples $SAMPLES (fresh_process=$process)" >>"$RESULTS/commands.txt"
        /usr/bin/time -v -o "$timing" "$BIN" \
            --fixture "$FIXTURE" \
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
    --extra "$HERE/run_targeted.sh" \
    --extra "$HERE/summarize_targeted.py" \
    --extra "$HERE/verify.py" \
    --extra "$FIXTURE"
cmp -s "$RESULTS/source-manifest-before.txt" "$RESULTS/source-manifest-after.txt"

{
    printf 'source_manifest_before_sha256='
    sha256sum "$RESULTS/source-manifest-before.txt" | cut -d' ' -f1
    printf 'source_manifest_after_sha256='
    sha256sum "$RESULTS/source-manifest-after.txt" | cut -d' ' -f1
    printf 'git_head='
    git -C "$ROOT" rev-parse HEAD
    printf 'git_status_relevant='
    git -C "$ROOT" status --short -- crates/litchi-drawingml/src/theme/family "$HERE" | tr '\n' '|'
    printf '\nfixture_sha256='
    sha256sum "$FIXTURE"
    printf '%s\n' 'source_sha256:'
    for path in \
        "$ROOT/crates/litchi-drawingml/src/theme/family/codec.rs" \
        "$ROOT/crates/litchi-drawingml/src/theme/family/mod.rs" \
        "$ROOT/crates/litchi-drawingml/src/theme/family/model.rs" \
        "$ROOT/crates/litchi-drawingml/src/theme/family/part.rs" \
        "$ROOT/crates/litchi-drawingml/src/theme/family/transaction.rs" \
        "$ROOT/crates/litchi-drawingml/tests/theme_part.rs"; do
        sha256sum "$path"
    done
    printf '%s\n' 'harness_sha256:'
    find "$HERE/harness" -type f -print0 | sort -z | xargs -0 sha256sum
} >"$RESULTS/source-provenance.txt"

python3 "$HERE/summarize_targeted.py" --results "$RESULTS" --output "$HERE/report.md"
python3 "$HERE/verify.py"
