#!/usr/bin/env bash
set -euo pipefail

# Current-source correctness gate. Historical run_smoke.sh and its d1 receipts
# remain unchanged; this runner writes only to a caller-supplied fresh path.
if [[ "${PROFILE_MODE:-smoke}" != smoke ]]; then
    echo "only the bounded smoke mode is available; full timing is gated" >&2
    exit 2
fi
for variable in RUSTFLAGS CARGO_ENCODED_RUSTFLAGS RUSTC_BOOTSTRAP RUSTDOCFLAGS; do
    if [[ -n "${!variable:-}" ]]; then
        echo "smoke runner refuses ${variable}" >&2
        exit 2
    fi
done
if [[ ! -x /usr/bin/time ]]; then
    echo "smoke runner requires /usr/bin/time -v for RSS evidence" >&2
    exit 2
fi

HERE=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd -P)
ROOT=$(git -C "$HERE/../../../.." rev-parse --show-toplevel)
SOURCE_COMMIT=d000d977b99e03f8542c7dae74acf767a91b1feb
HARNESS="$HERE/harness/Cargo.toml"
MANIFEST_TOOL="$HERE/source_manifest.py"
VERIFY_TOOL="$HERE/verify_current_smoke.py"
RESULTS_INPUT=${DOCX_STYLES_EFFECTS_RESULTS:-}
TARGET_INPUT=${DOCX_STYLES_EFFECTS_TARGET_DIR:-/var/tmp/litchi-docx-styles-effects-smoke-target}
if [[ -z "$RESULTS_INPUT" ]]; then
    echo "set DOCX_STYLES_EFFECTS_RESULTS to a fresh output directory" >&2
    exit 2
fi
RESULTS=$(realpath -m -- "$RESULTS_INPUT")
if [[ "$TARGET_INPUT" = /* ]]; then
    TARGET=$(realpath -m -- "$TARGET_INPUT")
else
    TARGET=$(realpath -m -- "$PWD/$TARGET_INPUT")
fi
if [[ "$RESULTS" == "/" || "$TARGET" == "/" ]]; then
    echo "refusing root as results or target" >&2
    exit 2
fi
if [[ "$RESULTS" == "$ROOT" || "$RESULTS" == "$ROOT/"* || "$ROOT" == "$RESULTS/"* ]]; then
    echo "results must be outside the clean source checkout: $RESULTS" >&2
    exit 2
fi
if [[ "$TARGET" == "$ROOT" || "$TARGET" == "$ROOT/"* || "$ROOT" == "$TARGET/"* ]]; then
    echo "Cargo target must be outside the clean source checkout: $TARGET" >&2
    exit 2
fi
if [[ "$RESULTS" == "$TARGET" || "$RESULTS" == "$TARGET/"* || "$TARGET" == "$RESULTS/"* ]]; then
    echo "results and Cargo target must be disjoint: $RESULTS / $TARGET" >&2
    exit 2
fi
if [[ -e "$RESULTS" || -L "$RESULTS" || -e "$TARGET" || -L "$TARGET" ]]; then
    echo "results and target must be fresh, absent paths" >&2
    exit 2
fi

CURRENT_COMMIT=$(git -C "$ROOT" --no-replace-objects rev-parse HEAD)
if ! git -C "$ROOT" --no-replace-objects merge-base --is-ancestor "$SOURCE_COMMIT" "$CURRENT_COMMIT"; then
    echo "current checkout does not descend from approved source $SOURCE_COMMIT" >&2
    exit 2
fi
if [[ -n "$(git -C "$ROOT" --no-replace-objects status --porcelain --untracked-files=all)" ]]; then
    echo "source checkout must be clean and committed before smoke" >&2
    exit 2
fi
for required in \
    "$HARNESS" "$HERE/harness/Cargo.lock" "$HERE/harness/adapter.rs" "$HERE/harness/main.rs" "$HERE/harness/support.rs" \
    "$HERE/source_manifest.py" "$HERE/test_source_snapshot.py" "$HERE/test_verify.py" "$HERE/verify.py" "$HERE/verify_current_smoke.py" "$HERE/run_current_smoke.sh" "$HERE/requirements.md" "$HERE/current-corpus-manifest.json" "$HERE/corpus-manifest.json" "$HERE/README.md" \
    "$HERE/run_smoke.sh" "$HERE/run_current_smoke.sh" "$HERE/verify_current_smoke.py" "$HERE/current-corpus-manifest.json" "$HERE/fixtures/Bug54849.docx" "$HERE/fixtures/ms-office-2010-signed.docx" \
    "$HERE/fixtures/ComplexNumberedLists.docx" "$HERE/fixtures/testGlossary.docx"; do
    [[ -f "$required" ]] || { echo "committed smoke input is missing: $required" >&2; exit 2; }
    git -C "$ROOT" --no-replace-objects ls-files --error-unmatch -- "$required" >/dev/null
done

LANES=(
    native_capture_bug_main native_capture_bug_glossary
    native_capture_signed_main native_capture_signed_glossary_absent
    native_capture_complex_main native_capture_complex_glossary_absent
    native_capture_glossary_main native_capture_glossary_glossary
    source_noop_main source_noop_glossary projection_main projection_glossary
    replace_main replace_glossary remove_main remove_glossary
    add_main_absent add_glossary_missing inverse_replace_main inverse_remove_main
    stale_patch_main signed_noop signed_changed independent_main independent_glossary
    cap_parts cap_total_part_bytes cap_total_relationships
    cap_total_relationship_xml_events cap_total_relationship_xml_bytes
    cap_relationship_parts cap_relationship_graph_nodes
    malformed_duplicate_owner malformed_third_orphan malformed_external
    malformed_wrong_content_type malformed_outbound malformed_shared_inbound
    malformed_root malformed_namespace malformed_opaque_xml
    malformed_unbound_descendant malformed_invalid_qname malformed_raw_attribute
    malformed_raw_text malformed_control malformed_invalid_char_ref
    malformed_empty_prefix malformed_reserved_xml_uri malformed_xml_version
    malformed_xml_events malformed_xml_depth
)

mkdir -p -- "$(dirname -- "$RESULTS")" "$(dirname -- "$TARGET")"
TARGET_CREATED=0
cleanup() {
    status=$?
    if [[ "$TARGET_CREATED" == 1 && -d "$TARGET" ]]; then
        find "$TARGET" -depth -delete
    fi
    exit "$status"
}
trap cleanup EXIT

mkdir -- "$RESULTS"
mkdir -- "$TARGET"
TARGET_CREATED=1

export CARGO_TARGET_DIR="$TARGET"
export CARGO_INCREMENTAL=0
export LC_ALL=C
unset RUSTFLAGS CARGO_ENCODED_RUSTFLAGS RUSTC_BOOTSTRAP RUSTDOCFLAGS

cargo metadata --format-version=1 --locked --offline --manifest-path "$HARNESS" \
    >"$RESULTS/smoke-metadata-before.json"

EXTRA_ARGS=(
    --extra "$ROOT/Cargo.toml"
    --extra "$ROOT/rust-toolchain.toml"
    --extra "$ROOT/.cargo/config.toml"
    --extra "$HERE/requirements.md"
    --extra "$HERE/current-corpus-manifest.json"
    --extra "$HERE/corpus-manifest.json"
    --extra "$HERE/README.md"
    --extra "$HERE/run_smoke.sh"
    --extra "$HERE/run_current_smoke.sh"
    --extra "$HERE/verify_current_smoke.py"
    --extra "$HERE/source_manifest.py"
    --extra "$HERE/test_source_snapshot.py"
    --extra "$HERE/test_verify.py"
    --extra "$HERE/verify.py"
    --extra "$HARNESS"
    --extra "$HERE/harness/Cargo.lock"
    --extra "$HERE/harness/adapter.rs"
    --extra "$HERE/harness/main.rs"
    --extra "$HERE/harness/support.rs"
)
for source in \
    "$ROOT/crates/litchi-docx/src/styles/effects.rs" \
    "$ROOT/crates/litchi-docx/src/package/package/styles_with_effects.rs" \
    "$ROOT/crates/litchi-docx/tests/styles_with_effects.rs" \
    "$ROOT/crates/litchi-opc/src/phys_pkg.rs" "$ROOT/crates/litchi-opc/src/limits.rs"; do
    EXTRA_ARGS+=(--extra "$source")
done
for fixture in "$HERE"/fixtures/*.docx; do
    EXTRA_ARGS+=(--extra "$fixture")
done
python3 "$MANIFEST_TOOL" \
    --metadata "$RESULTS/smoke-metadata-before.json" \
    --root "$ROOT" --evidence "$HERE" \
    --output "$RESULTS/smoke-source-manifest-before.txt" \
    --source-commit "$SOURCE_COMMIT" "${EXTRA_ARGS[@]}"

cargo build --locked --offline --manifest-path "$HARNESS" \
    >"$RESULTS/smoke-build.log" 2>&1
BIN="$TARGET/debug/docx-styles-effects-smoke"
[[ -x "$BIN" ]] || { echo "smoke binary was not produced: $BIN" >&2; exit 1; }
sha256sum "$BIN" >"$RESULTS/smoke-binary.sha256"
{
    printf 'binary=%s\n' "$BIN"
    cat "$RESULTS/smoke-binary.sha256"
    printf 'source_commit=%s\n' "$SOURCE_COMMIT"
    printf 'git_head=%s\n' "$CURRENT_COMMIT"
    printf 'git_status=clean\n'
    printf 'rustc=%s\n' "$(rustc -vV | tr '\n' ' ')"
    printf 'cargo=%s\n' "$(cargo -V)"
    printf 'target=%s\n' "$TARGET"
    printf 'allocator=CountingAllocator (process-local GlobalAlloc observer)\n'
    printf 'rss=/usr/bin/time -v Maximum resident set size\n'
    printf 'mode=smoke processes=1 warmup=0 samples=1\n'
} >"$RESULTS/smoke-build-provenance.txt"

: >"$RESULTS/smoke-commands.txt"
{
    printf 'command=cargo metadata --format-version=1 --locked --offline --manifest-path %q\n' "$HARNESS"
    printf 'command=python3 %q --metadata %q --root %q --evidence %q --output %q --source-commit %q\n' \
        "$MANIFEST_TOOL" "$RESULTS/smoke-metadata-before.json" "$ROOT" "$HERE" \
        "$RESULTS/smoke-source-manifest-before.txt" "$SOURCE_COMMIT"
    printf 'command=cargo build --locked --offline --manifest-path %q\n' "$HARNESS"
    printf 'environment=CARGO_TARGET_DIR=%q CARGO_INCREMENTAL=0 LC_ALL=%q RUSTFLAGS=unset CARGO_ENCODED_RUSTFLAGS=unset RUSTC_BOOTSTRAP=unset RUSTDOCFLAGS=unset\n' \
        "$TARGET" "$LC_ALL"
} >>"$RESULTS/smoke-commands.txt"
for lane in "${LANES[@]}"; do
    report="$RESULTS/smoke-${lane}-p1.json"
    timing="$RESULTS/smoke-${lane}-p1.time.txt"
    stderr="$RESULTS/smoke-${lane}-p1.stderr.log"
    printf 'run=/usr/bin/time -v %q --lane %q --warmup 0 --samples 1 (fresh_process=1)\n' \
        "$BIN" "$lane" >>"$RESULTS/smoke-commands.txt"
    /usr/bin/time -v -o "$timing" "$BIN" --lane "$lane" --warmup 0 --samples 1 \
        >"$report" 2>"$stderr"
done

cargo metadata --format-version=1 --locked --offline --manifest-path "$HARNESS" \
    >"$RESULTS/smoke-metadata-after.json"
python3 "$MANIFEST_TOOL" \
    --metadata "$RESULTS/smoke-metadata-after.json" \
    --root "$ROOT" --evidence "$HERE" \
    --output "$RESULTS/smoke-source-manifest-after.txt" \
    --source-commit "$SOURCE_COMMIT" "${EXTRA_ARGS[@]}"
cmp -s "$RESULTS/smoke-source-manifest-before.txt" "$RESULTS/smoke-source-manifest-after.txt"
sha256sum "$BIN" >"$RESULTS/smoke-binary-after.sha256"
cmp -s "$RESULTS/smoke-binary.sha256" "$RESULTS/smoke-binary-after.sha256"
{
    printf 'source_manifest_sha256='
    sha256sum "$RESULTS/smoke-source-manifest-before.txt" | cut -d' ' -f1
    printf 'metadata_before_sha256='
    sha256sum "$RESULTS/smoke-metadata-before.json" | cut -d' ' -f1
    printf 'metadata_after_sha256='
    sha256sum "$RESULTS/smoke-metadata-after.json" | cut -d' ' -f1
    printf 'git_head=%s\n' "$CURRENT_COMMIT"
} >"$RESULTS/smoke-source-provenance.txt"

python3 "$VERIFY_TOOL" \
    --root "$ROOT" --evidence "$HERE" --results "$RESULTS" \
    --manifest "$RESULTS/smoke-source-manifest-before.txt" \
    --manifest-after "$RESULTS/smoke-source-manifest-after.txt" \
    --metadata-before "$RESULTS/smoke-metadata-before.json" \
    --metadata-after "$RESULTS/smoke-metadata-after.json" \
    --corpus "$HERE/current-corpus-manifest.json" \
    --output "$RESULTS/smoke-verification.json"
echo "stylesWithEffects smoke passed; receipts: $RESULTS"
