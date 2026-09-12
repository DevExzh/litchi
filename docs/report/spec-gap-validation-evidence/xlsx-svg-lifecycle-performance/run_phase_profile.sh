#!/usr/bin/env bash
set -euo pipefail

if [[ "${PROFILE_FROZEN:-}" != "1" ]]; then
    echo "refusing phase profile without PROFILE_FROZEN=1" >&2
    exit 2
fi
if [[ "${XLSX_SVG_PROFILE_API_WIRED:-}" != "1" ]]; then
    echo "refusing phase profile before API wiring gate" >&2
    exit 2
fi
for variable in RUSTFLAGS CARGO_ENCODED_RUSTFLAGS RUSTC_BOOTSTRAP; do
    if [[ -n "${!variable:-}" ]]; then
        echo "refusing phase profile with ${variable} set" >&2
        exit 2
    fi
done
if [[ ! -x /usr/bin/time ]]; then
    echo "refusing phase profile because /usr/bin/time -v is unavailable" >&2
    exit 2
fi

HERE=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd -P)
ROOT=$(cd -- "$HERE/../../../.." && pwd -P)
HARNESS="$HERE/harness/Cargo.toml"
MANIFEST_TOOL="$HERE/source_manifest.py"
PROFILE_PINS="$HERE/profile_pins.py"
COMMITTED_INPUTS="$HERE/committed_inputs.py"

if [[ -z "${XLSX_SVG_PHASE_RESULTS:-}" ]]; then
    echo "set XLSX_SVG_PHASE_RESULTS to a fresh external output directory" >&2
    exit 2
fi
RESULTS=$(realpath -m -- "$XLSX_SVG_PHASE_RESULTS")
TARGET_INPUT=${CARGO_TARGET_DIR:-/var/tmp/litchi-xlsx-svg-phase-target}
if [[ "$TARGET_INPUT" = /* ]]; then
    TARGET=$(realpath -m -- "$TARGET_INPUT")
else
    TARGET=$(realpath -m -- "$PWD/$TARGET_INPUT")
fi
if [[ "$RESULTS" == "$ROOT" || "$RESULTS" == "$ROOT/"* || "$ROOT" == "$RESULTS/"* ]]; then
    echo "phase results must be outside the checkout: $RESULTS" >&2
    exit 2
fi
if [[ "$TARGET" == "$ROOT" || "$TARGET" == "$ROOT/"* || "$ROOT" == "$TARGET/"* || "$TARGET" == "/" ]]; then
    echo "phase target must be outside the checkout: $TARGET" >&2
    exit 2
fi
if [[ "$RESULTS" == "/" || "$RESULTS" == "$TARGET" || "$RESULTS" == "$TARGET/"* || "$TARGET" == "$RESULTS/"* ]]; then
    echo "phase results and target must be disjoint" >&2
    exit 2
fi
if [[ -e "$RESULTS" || -L "$RESULTS" || -L "$XLSX_SVG_PHASE_RESULTS" ]]; then
    echo "refusing existing phase output: $RESULTS" >&2
    exit 2
fi
if [[ -e "$TARGET" || -L "$TARGET" ]]; then
    echo "refusing existing phase target; choose a fresh target: $TARGET" >&2
    exit 2
fi

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
    echo "phase evidence requires 3 processes, at least 2 warmups, and at least 20 samples" >&2
    exit 2
fi

TARGET_CREATED=0
cleanup_target() {
    if [[ "$TARGET_CREATED" == "1" ]]; then
        find "$TARGET" -depth -delete
    fi
}
trap cleanup_target EXIT

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
    echo "requested source pin differs from the committed profile pin" >&2
    exit 2
fi
if ! git -C "$ROOT" merge-base --is-ancestor "$SOURCE_PIN" "$CURRENT_COMMIT"; then
    echo "approved source pin is not an ancestor of the phase checkout: $SOURCE_PIN" >&2
    exit 2
fi
mapfile -t PINNED_SOURCES < <(python3 "$PROFILE_PINS" --paths)
guard_args=(--root "$ROOT" --commit "$SOURCE_PIN")
for source in "${PINNED_SOURCES[@]}"; do
    guard_args+=(--path "$source")
done
python3 "$COMMITTED_INPUTS" "${guard_args[@]}"

manifest_args=(
    --metadata "$RESULTS/metadata-before.json"
    --root "$ROOT"
    --output "$RESULTS/source-manifest-before.txt"
    --extra "$ROOT/Cargo.toml"
    --extra "$ROOT/rust-toolchain.toml"
    --extra "$ROOT/.cargo/config.toml"
    --extra "$HARNESS"
    --extra "$HERE/harness/Cargo.lock"
    --extra "$HERE/harness/adapter.rs"
    --extra "$HERE/harness/support.rs"
    --extra "$HERE/harness/main.rs"
    --extra "$HERE/harness/phase_main.rs"
    --extra "$MANIFEST_TOOL"
    --extra "$PROFILE_PINS"
    --extra "$COMMITTED_INPUTS"
    --extra "$HERE/run_phase_profile.sh"
    --extra "$HERE/README.md"
    --extra "$HERE/requirements.md"
    --extra "$HERE/root-review.md"
    --extra "$HERE/phase_verify.py"
    --extra "$HERE/phase_summarize.py"
    --extra "$HERE/test_phase_profile.py"
    --extra "$HERE/phase-decomposition.md"
    --extra "$HERE/corpus-manifest.json"
    --extra "$HERE/fixtures/tdf169496_hidden_graphic.xlsx"
    --extra "$ROOT/docs/adr/0001-priorities-and-api-layers.md"
    --extra "$ROOT/docs/adr/0005-io-memory-and-performance.md"
    --extra "$ROOT/docs/report/spec-gap-validation-evidence/xlsx-svg-lifecycle-design.md"
    --git-commit "$CURRENT_COMMIT"
)
python3 "$MANIFEST_TOOL" "${manifest_args[@]}"

cargo build --release --locked --offline --manifest-path "$HARNESS" \
    --bin xlsx-svg-lifecycle-phase-profile >"$RESULTS/build.log" 2>&1
BIN="$TARGET/release/xlsx-svg-lifecycle-phase-profile"
if [[ ! -x "$BIN" ]]; then
    echo "phase profile binary was not produced: $BIN" >&2
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
    printf '%s\n' 'compiler_flags=RUSTFLAGS=unset CARGO_ENCODED_RUSTFLAGS=unset RUSTC_BOOTSTRAP=unset'
    printf '%s\n' 'allocator=CountingAllocator (process-local GlobalAlloc observer)'
    printf '%s\n' 'phase_scope=open,stages,commit,firstsave,reopen_secondsave,validation'
    printf '%s\n' "git_head=$CURRENT_COMMIT"
    printf '%s\n' "approved_source_pin=$SOURCE_PIN"
    printf '%s\n' "os=$(uname -srm 2>/dev/null || printf unavailable)"
    printf '%s\n' "cpu_model=$(awk -F: '/model name|Hardware/ {gsub(/^ +/, "", $2); print $2; exit}' /proc/cpuinfo 2>/dev/null || printf unavailable)"
    printf '%s\n' "core_count=$(nproc 2>/dev/null || printf unavailable)"
    printf '%s\n' "memory_total=$(awk '/MemTotal:/ {print $2 " " $3; exit}' /proc/meminfo 2>/dev/null || printf unavailable)"
    printf '%s\n' "storage=$(df -P "$ROOT" 2>/dev/null | tail -n 1 || printf unavailable)"
    printf '%s\n' "environment=CARGO_TARGET_DIR=$CARGO_TARGET_DIR CARGO_INCREMENTAL=$CARGO_INCREMENTAL"
} >"$RESULTS/build-provenance.txt"

: >"$RESULTS/commands.txt"
for pictures in 16 64 256; do
    process=1
    while [[ "$process" -le "$PROCESSES" ]]; do
        report="$RESULTS/phase_${pictures}-p${process}.json"
        timing="$RESULTS/phase_${pictures}-p${process}.time.txt"
        printf '%s\n' \
            "run=/usr/bin/time -v $BIN --pictures $pictures --warmup $WARMUP --samples $SAMPLES (fresh_process=$process)" \
            >>"$RESULTS/commands.txt"
        /usr/bin/time -v -o "$timing" "$BIN" \
            --pictures "$pictures" --warmup "$WARMUP" --samples "$SAMPLES" \
            >"$report" 2>"$RESULTS/phase_${pictures}-p${process}.stderr.log"
        process=$((process + 1))
    done
done

sha256sum "$BIN" >"$RESULTS/binary-after.sha256"
cmp -s "$RESULTS/binary.sha256" "$RESULTS/binary-after.sha256"
cargo metadata --format-version=1 --locked --offline --manifest-path "$HARNESS" \
    >"$RESULTS/metadata-after.json"
manifest_args[1]="$RESULTS/metadata-after.json"
manifest_args[5]="$RESULTS/source-manifest-after.txt"
python3 "$MANIFEST_TOOL" "${manifest_args[@]}"
cmp -s "$RESULTS/source-manifest-before.txt" "$RESULTS/source-manifest-after.txt"

{
    printf 'source_manifest_before_sha256='
    sha256sum "$RESULTS/source-manifest-before.txt" | cut -d' ' -f1
    printf 'source_manifest_after_sha256='
    sha256sum "$RESULTS/source-manifest-after.txt" | cut -d' ' -f1
    printf 'git_head=%s\n' "$CURRENT_COMMIT"
    printf 'approved_source_pin=%s\n' "$SOURCE_PIN"
} >"$RESULTS/source-provenance.txt"

python3 "$HERE/phase_summarize.py" --results "$RESULTS" --output "$RESULTS/report.md"
python3 "$HERE/phase_verify.py" --results "$RESULTS" --report "$RESULTS/report.md" \
    --output "$RESULTS/verification.json"
