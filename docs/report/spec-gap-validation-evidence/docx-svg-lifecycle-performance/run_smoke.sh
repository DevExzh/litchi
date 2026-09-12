#!/usr/bin/env bash
set -euo pipefail

HERE=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
ROOT=$(cd -- "$HERE/../../../.." && pwd)
BASELINE=892441d95db29da4351390716ef5c65b4c7c97de
COMMITTED_ROOT=${COMMITTED_ROOT:-/tmp/litchi-docx-svg-lifecycle-892441d95}
RUN_EVIDENCE="$COMMITTED_ROOT/docs/report/spec-gap-validation-evidence/docx-svg-lifecycle-performance"
RESULTS="$HERE/results"
MODE=${PROFILE_MODE:-smoke}

if [[ "$MODE" == full ]]; then
    PROCESSES=${PROCESSES:-3}
    WARMUP=${WARMUP:-2}
    SAMPLES=${SAMPLES:-20}
else
    PROCESSES=${PROCESSES:-1}
    WARMUP=${WARMUP:-0}
    SAMPLES=${SAMPLES:-1}
fi
if [[ "$MODE" != smoke && "$MODE" != full ]]; then
    echo "unknown profile mode: $MODE" >&2
    exit 2
fi
REPORT="$HERE/${MODE}-report.md"
DOMINANT="$HERE/${MODE}-dominant-costs.md"
VERIFICATION="$HERE/${MODE}-verification.json"
LANES=(
    native_svg_capture native_floating_capture
    lazy_inventory_1 lazy_inventory_64
    single_attach_1 single_attach_16 single_attach_64
    single_detach_1 single_detach_16 single_detach_64
    batch_attach_1 batch_attach_16 batch_attach_64
    batch_detach_1 batch_detach_16 batch_detach_64
    shared_svg_cleanup exact_inverse_single_1 exact_inverse_batch_64
    large_unchanged_media_managed_cap noop_detach_64
)
if [[ ! -d "$COMMITTED_ROOT/.git" && ! -f "$COMMITTED_ROOT/.git" ]]; then
    echo "COMMITTED_ROOT is not a Git checkout: $COMMITTED_ROOT" >&2
    exit 2
fi
if [[ "$(git -C "$COMMITTED_ROOT" rev-parse HEAD)" != "$BASELINE" ]]; then
    echo "COMMITTED_ROOT is not baseline $BASELINE" >&2
    exit 2
fi
if [[ -n "$(git -C "$COMMITTED_ROOT" status --porcelain)" ]]; then
    echo "COMMITTED_ROOT must be clean before profile staging" >&2
    exit 2
fi
if [[ -e "$RUN_EVIDENCE" ]]; then
    echo "refusing to overwrite existing isolated evidence path: $RUN_EVIDENCE" >&2
    exit 2
fi
mkdir -p -- "$(dirname "$RUN_EVIDENCE")" "$RUN_EVIDENCE/harness" "$RUN_EVIDENCE/fixtures"
STAGED=1
TARGET_CREATED=0
TARGET=""
cleanup_staged() {
    if [[ "${STAGED:-0}" == 1 ]]; then
        rm -rf -- "$RUN_EVIDENCE"
    fi
}
cleanup_target() {
    if [[ "${TARGET_CREATED:-0}" == 1 && -n "${TARGET:-}" ]]; then
        rm -rf -- "$TARGET"
    fi
}
cleanup() {
    cleanup_target
    cleanup_staged
}
trap cleanup EXIT

if [[ ! -d "$HERE/baseline-source" ]]; then
    echo "retained baseline source snapshot is missing: $HERE/baseline-source" >&2
    exit 2
fi
mkdir -p -- "$RUN_EVIDENCE/baseline-source"
cp -pR -- "$HERE/baseline-source/." "$RUN_EVIDENCE/baseline-source/"

for path in "$HERE"/harness/*.rs "$HERE"/harness/Cargo.toml "$HERE"/harness/Cargo.lock; do
    cp -p -- "$path" "$RUN_EVIDENCE/harness/"
done
for path in "$HERE"/fixtures/*.docx; do
    cp -p -- "$path" "$RUN_EVIDENCE/fixtures/"
done
for path in "$HERE"/*.md "$HERE"/*.json "$HERE"/*.py "$HERE"/*.sh; do
    cp -p -- "$path" "$RUN_EVIDENCE/"
done

if [[ -n "${CARGO_TARGET_DIR:-}" ]]; then
    TARGET=$CARGO_TARGET_DIR
    if [[ -e "$TARGET" && "${ALLOW_EXISTING_TARGET:-}" != 1 ]]; then
        echo "refusing to reuse existing Cargo target: $TARGET" >&2
        exit 2
    fi
    if [[ ! -e "$TARGET" ]]; then
        mkdir -p -- "$TARGET"
        TARGET_CREATED=1
    fi
else
    TARGET=$(mktemp -d /var/tmp/litchi-docx-svg-lifecycle-target.XXXXXX)
    TARGET_CREATED=1
fi

mkdir -p -- "$RESULTS"
# Remove only artifacts owned by this mode.  In particular, a smoke run must
# preserve full-profile receipts and unrelated evidence in this directory.
MODE_RESULTS=(
    "${MODE}-commands.txt"
    "${MODE}-source-manifest-before.txt"
    "${MODE}-source-manifest-after.txt"
    "${MODE}-source-provenance.txt"
    "${MODE}-build-provenance.txt"
    "${MODE}-build.log"
    "${MODE}-binary.sha256"
    "${MODE}-binary-after.sha256"
    "${MODE}-metadata-before.json"
    "${MODE}-metadata-after.json"
)
for name in "${MODE_RESULTS[@]}"; do
    rm -f -- "$RESULTS/$name"
done
for lane in "${LANES[@]}"; do
    for process in 1 2 3; do
        prefix="${MODE}-${lane}-p${process}"
        rm -f -- "$RESULTS/${prefix}.json" \
            "$RESULTS/${prefix}.time.txt" "$RESULTS/${prefix}.stderr.log"
    done
done
rm -f -- "$REPORT" "$DOMINANT" "$VERIFICATION"

unset RUSTFLAGS CARGO_ENCODED_RUSTFLAGS RUSTC_BOOTSTRAP
export CARGO_TARGET_DIR="$TARGET"
export CARGO_INCREMENTAL=0
export LC_ALL=C
HARNESS="$RUN_EVIDENCE/harness/Cargo.toml"

cargo metadata --format-version=1 --locked --offline --manifest-path "$HARNESS" \
    >"$RESULTS/${MODE}-metadata-before.json"
python3 "$HERE/source_manifest.py" --root "$COMMITTED_ROOT" --evidence "$RUN_EVIDENCE" \
    --output "$RESULTS/${MODE}-source-manifest-before.txt"
cargo build --release --locked --offline --manifest-path "$HARNESS" \
    >"$RESULTS/${MODE}-build.log" 2>&1
BIN="$TARGET/release/docx-svg-lifecycle-profile"
[[ -x "$BIN" ]] || { echo "profile binary was not produced: $BIN" >&2; exit 1; }
sha256sum "$BIN" >"$RESULTS/${MODE}-binary.sha256"
{
    printf 'binary=%s\n' "$BIN"
    cat "$RESULTS/${MODE}-binary.sha256"
    printf 'git_head=%s\n' "$(git -C "$COMMITTED_ROOT" rev-parse HEAD)"
    printf 'rustc=%s\n' "$(rustc -vV | tr '\n' ' ')"
    printf 'cargo=%s\n' "$(cargo -V)"
    printf 'target=%s\n' "$TARGET"
    printf 'allocator=CountingAllocator (process-local GlobalAlloc observer)\n'
    printf 'rss=/usr/bin/time -v Maximum resident set size\n'
    printf 'mode=%s processes=%s warmup=%s samples=%s\n' "$MODE" "$PROCESSES" "$WARMUP" "$SAMPLES"
} >"$RESULTS/${MODE}-build-provenance.txt"

: >"$RESULTS/${MODE}-commands.txt"
for lane in "${LANES[@]}"; do
    for process in $(seq 1 "$PROCESSES"); do
        prefix="$MODE-${lane}-p${process}"
        report="$RESULTS/${prefix}.json"
        timing="$RESULTS/${prefix}.time.txt"
        stderr="$RESULTS/${prefix}.stderr.log"
        printf '%s\n' "run=/usr/bin/time -v $BIN --lane $lane --warmup $WARMUP --samples $SAMPLES (fresh_process=$process)" \
            >>"$RESULTS/${MODE}-commands.txt"
        /usr/bin/time -v -o "$timing" "$BIN" --lane "$lane" \
            --warmup "$WARMUP" --samples "$SAMPLES" >"$report" 2>"$stderr"
    done
done

cargo metadata --format-version=1 --locked --offline --manifest-path "$HARNESS" \
    >"$RESULTS/${MODE}-metadata-after.json"
python3 "$HERE/source_manifest.py" --root "$COMMITTED_ROOT" --evidence "$RUN_EVIDENCE" \
    --output "$RESULTS/${MODE}-source-manifest-after.txt"
cmp -s "$RESULTS/${MODE}-source-manifest-before.txt" "$RESULTS/${MODE}-source-manifest-after.txt"
sha256sum "$BIN" >"$RESULTS/${MODE}-binary-after.sha256"
cmp -s "$RESULTS/${MODE}-binary.sha256" "$RESULTS/${MODE}-binary-after.sha256"

{
    printf 'source_manifest_sha256='
    sha256sum "$RESULTS/${MODE}-source-manifest-before.txt" | cut -d' ' -f1
    printf 'git_head='
    git -C "$COMMITTED_ROOT" rev-parse HEAD
    printf '%s\n' 'source_manifest:'
    cat "$RESULTS/${MODE}-source-manifest-before.txt"
} >"$RESULTS/${MODE}-source-provenance.txt"

if [[ "$MODE" == smoke ]]; then
    python3 "$HERE/summarize.py" --results "$RESULTS" --mode smoke --output "$REPORT" \
        --dominant-output "$DOMINANT"
else
    python3 "$HERE/summarize.py" --results "$RESULTS" --mode full --output "$REPORT" \
        --dominant-output "$DOMINANT"
fi
python3 "$HERE/verify.py" --evidence "$HERE" --native-root "$ROOT" --results "$RESULTS" \
    --manifest "$RESULTS/${MODE}-source-manifest-before.txt" --mode "$MODE" \
    --report "$REPORT" --corpus "$HERE/corpus-manifest.json" \
    --output "$VERIFICATION"
