#!/usr/bin/env bash
set -euo pipefail

if [[ "${PROFILE_FROZEN:-}" != "1" ]]; then
    echo "refusing to profile an unfrozen XLSX source snapshot; set PROFILE_FROZEN=1" >&2
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
# contextual-performance is one directory below the lifecycle evidence root.
ROOT=$(cd -- "$HERE/../../../../.." && pwd)
PROFILE_LABEL=${PROFILE_LABEL:-}
case "$PROFILE_LABEL" in
    before-8b5838c59)
        EXPECTED_COMMIT=8b5838c591775990747b2cbce82fb2eea372b58b
        EXPECTED_SOURCE_HASH=1d8e6adf127a00dc25cd801bd90835f41442220ce944ba005cfdfa28e6dd62c2
        EXPECTED_TEST_HASH=03c67af436e218156cd48df52de1963bff97006025bc9f2d1848122f910bcc89
        EXPECTED_CODEC_HASH=e8a1e480d249b834f286e0deec083f8044257faf3b95905fb5b771e83f71605a
        ;;
    after-fc44c4e6c)
        EXPECTED_COMMIT=fc44c4e6c945ab07ded7447f40670d898839eeb3
        EXPECTED_SOURCE_HASH=952cbc65b89314c61ecee623be91af64665d3c92b6827a58fbceaf8a722efb00
        EXPECTED_TEST_HASH=7278bc01cf6bb84234e2f359064a3cf85cff6a8f270ec8ea9b7a83d868c9e34e
        EXPECTED_CODEC_HASH=e8a1e480d249b834f286e0deec083f8044257faf3b95905fb5b771e83f71605a
        ;;
    *) echo "PROFILE_LABEL must identify the frozen before or after snapshot" >&2; exit 2 ;;
esac
ACTUAL_COMMIT=$(git -C "$ROOT" rev-parse HEAD)
if [[ "$ACTUAL_COMMIT" != "$EXPECTED_COMMIT" ]]; then
    echo "frozen commit mismatch: expected $EXPECTED_COMMIT, got $ACTUAL_COMMIT" >&2
    exit 2
fi
if [[ -n "$(git -C "$ROOT" status --porcelain -- crates)" ]]; then
    echo "refusing dirty production source tree under crates/" >&2
    git -C "$ROOT" status --short -- crates >&2
    exit 2
fi
HARNESS="$HERE/harness/Cargo.toml"
RESULTS="$HERE/results"
TARGET_INPUT=${CARGO_TARGET_DIR:-/var/tmp/litchi-xlsx-contextual-profile-target}
if [[ "$TARGET_INPUT" = /* ]]; then
    TARGET=$(realpath -m -- "$TARGET_INPUT")
else
    TARGET=$(realpath -m -- "$PWD/$TARGET_INPUT")
fi
MANIFEST_TOOL="$HERE/source_manifest.py"
LANES=(
    contextual_read_p16_n0 contextual_read_p16_n32 contextual_read_p16_n128
    contextual_read_p32_n0 contextual_read_p32_n32 contextual_read_p32_n128
    contextual_read_p128_n0 contextual_read_p128_n32 contextual_read_p128_n128
    standalone_export_p16_n0 standalone_export_p16_n32 standalone_export_p16_n128
    standalone_export_p32_n0 standalone_export_p32_n32 standalone_export_p32_n128
    standalone_export_p128_n0 standalone_export_p128_n32 standalone_export_p128_n128
    scalar_reference_edit_p16_n0 scalar_reference_edit_p16_n32 scalar_reference_edit_p16_n128
    scalar_reference_edit_p32_n0 scalar_reference_edit_p32_n32 scalar_reference_edit_p32_n128
    scalar_reference_edit_p128_n0 scalar_reference_edit_p128_n32 scalar_reference_edit_p128_n128
    small_cap_refusal_p16_n0 small_cap_refusal_p16_n32 small_cap_refusal_p16_n128
    small_cap_refusal_p32_n0 small_cap_refusal_p32_n32 small_cap_refusal_p32_n128
    small_cap_refusal_p128_n0 small_cap_refusal_p128_n32 small_cap_refusal_p128_n128
    contextual_read_original_149433 standalone_export_original_149433
    small_cap_refusal_original_149433
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
if [[ -e "$TARGET" ]]; then
    echo "refusing to reuse an existing target; choose a fresh caller-owned path: $TARGET" >&2
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
rm -f "$RESULTS"/*.json "$RESULTS"/*.time.txt "$RESULTS"/*.stderr.txt "$RESULTS"/*.log \
    "$RESULTS"/commands.txt "$RESULTS"/source-manifest-*.txt \
    "$RESULTS"/source-provenance.txt "$RESULTS"/binary.sha256 "$RESULTS"/corpus-sha256.tsv \
    "$RESULTS"/build-provenance.txt "$RESULTS"/provenance-verification.json \
    "$HERE/report.md" "$HERE/verification.json"

unset RUSTFLAGS CARGO_ENCODED_RUSTFLAGS RUSTC_BOOTSTRAP
export CARGO_TARGET_DIR="$TARGET"
export CARGO_INCREMENTAL=0

cargo metadata --format-version=1 --locked --offline --manifest-path "$HARNESS" \
    >"$RESULTS/metadata-before.json"
manifest_extras=(
    --extra "$ROOT/Cargo.toml"
    --extra "$ROOT/Cargo.lock"
    --extra "$HARNESS"
    --extra "$HERE/harness/Cargo.lock"
    --extra "$HERE/harness/main.rs"
    --extra "$HERE/harness/support.rs"
    --extra "$MANIFEST_TOOL"
    --extra "$HERE/validate_provenance.py"
    --extra "$HERE/corpus_hashes.py"
    --extra "$HERE/compare_manifests.py"
    --extra "$HERE/run_profile.sh"
    --extra "$HERE/summarize.py"
    --extra "$HERE/verify.py"
    --extra "$HERE/README.md"
    --extra "$HERE/requirements.md"
    --extra "$HERE/corpus-manifest.json"
    --extra "$ROOT/docs/GOAL.md"
    --extra "$ROOT/docs/report/spec-gap-validation-evidence/xlsx-svg-lifecycle-design.md"
)
python3 "$MANIFEST_TOOL" \
    --metadata "$RESULTS/metadata-before.json" --root "$ROOT" \
    --output "$RESULTS/source-manifest-before.txt" "${manifest_extras[@]}"

cargo build --release --locked --offline --manifest-path "$HARNESS" \
    >"$RESULTS/build.log" 2>&1
BIN="$TARGET/release/xlsx-svg-contextual-profile"
if [[ ! -x "$BIN" ]]; then
    echo "profile binary was not produced: $BIN" >&2
    exit 1
fi
sha256sum "$BIN" >"$RESULTS/binary.sha256"
cpu_model=$(awk -F: '/model name|Hardware/ {gsub(/^ +/, "", $2); print $2; exit}' /proc/cpuinfo 2>/dev/null || printf unavailable)
memory_total=$(awk '/MemTotal:/ {print $2 " " $3; exit}' /proc/meminfo 2>/dev/null || printf unavailable)
storage=$(df -P "$ROOT" 2>/dev/null | tail -n 1 || printf unavailable)
{
    printf 'snapshot_label=%s\n' "${PROFILE_LABEL:-unspecified}"
    printf 'git_head='; git -C "$ROOT" rev-parse HEAD
    printf 'binary=%s\n' "$BIN"
    cat "$RESULTS/binary.sha256"
    printf '%s\n' 'rustc -vV:'; rustc -vV
    printf '%s\n' "cargo=$(cargo -V)"
    printf '%s\n' "target=$TARGET"
    printf '%s\n' 'fresh_target=1'
    printf '%s\n' "cargo_incremental=$CARGO_INCREMENTAL"
    printf '%s\n' 'allocator=CountingAllocator (process-local GlobalAlloc observer)'
    printf '%s\n' 'source_api=SourceDrawing::scan + SvgOwner::value + NamespaceContext::shares_storage'
    printf '%s\n' "os=$(uname -srm 2>/dev/null || printf unavailable)"
    printf '%s\n' "cpu_model=$cpu_model"
    printf '%s\n' "core_count=$(nproc 2>/dev/null || printf unavailable)"
    printf '%s\n' "memory_total=$memory_total"
    printf '%s\n' "storage=$storage"
    printf '%s\n' "environment=CARGO_TARGET_DIR=$CARGO_TARGET_DIR CARGO_INCREMENTAL=$CARGO_INCREMENTAL"
    printf '%s\n' 'source_sha256:'
    for path in \
        "$ROOT/crates/litchi-xlsx/src/drawing/source.rs" \
        "$ROOT/crates/litchi-xlsx/src/drawing/source_tests.rs" \
        "$ROOT/crates/litchi-drawingml/src/svg_blip.rs"; do
        sha256sum "$path"
    done
} >"$RESULTS/build-provenance.txt"
python3 "$HERE/validate_provenance.py" \
    --results "$RESULTS" --binary "$BIN" \
    --expected-commit "$EXPECTED_COMMIT" --expected-label "$PROFILE_LABEL" \
    --expected-source-hash "$EXPECTED_SOURCE_HASH" --expected-test-hash "$EXPECTED_TEST_HASH" \
    --expected-codec-hash "$EXPECTED_CODEC_HASH"

: >"$RESULTS/commands.txt"
for lane in "${LANES[@]}"; do
    process=1
    while [[ "$process" -le "$PROCESSES" ]]; do
        report="$RESULTS/${lane}-p${process}.json"
        timing="$RESULTS/${lane}-p${process}.time.txt"
        stderr="$RESULTS/${lane}-p${process}.stderr.txt"
        printf '%s\n' \
            "run=/usr/bin/time -v -o $timing $BIN --lane $lane --warmup $WARMUP --samples $SAMPLES > $report 2> $stderr (fresh_process=$process)" \
            >>"$RESULTS/commands.txt"
        /usr/bin/time -v -o "$timing" "$BIN" \
            --lane "$lane" --warmup "$WARMUP" --samples "$SAMPLES" \
            >"$report" 2>"$stderr"
        process=$((process + 1))
    done
done

cargo metadata --format-version=1 --locked --offline --manifest-path "$HARNESS" \
    >"$RESULTS/metadata-after.json"
python3 "$MANIFEST_TOOL" \
    --metadata "$RESULTS/metadata-after.json" --root "$ROOT" \
    --output "$RESULTS/source-manifest-after.txt" "${manifest_extras[@]}"
cmp -s "$RESULTS/source-manifest-before.txt" "$RESULTS/source-manifest-after.txt"

python3 "$HERE/corpus_hashes.py" --results "$RESULTS" --output "$RESULTS/corpus-sha256.tsv"
python3 "$HERE/summarize.py" --results "$RESULTS" --output "$HERE/report.md"
python3 "$HERE/verify.py" --snapshot "$PROFILE_LABEL" --expected-commit "$EXPECTED_COMMIT"
