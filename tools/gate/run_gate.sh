#!/bin/bash
# Coordinator gate: the 53 rows, run serially, one summary file.
#
# Usage: tools/gate/run_gate.sh <name> [row ...]
#   <name>   result directory name.
#   row      optional selection by index (41) or name (fuzz); default all rows.
#   --list   print the row table and exit.
#
# Environment:
#   GATE_DIR       checkout to gate (default: the checkout holding this script).
#   GATE_OUT       result directory (default: $GATE_DIR/target/gate-results/<name>).
#   GATE_EXTRA     extra clippy `-p` arguments for row 1.
#   GATE_OFFLINE   set to 1 to forbid network access (CARGO_NET_OFFLINE).
#   GATE_LOCK_SEED Cargo.lock to copy into a checkout that has none (the root
#                  lock is untracked by policy, ADR 0028).
#   HARNESS_TESTS  set to run the perf-harness tests instead of `cargo check`.
#   CARGO_HOME, RUSTUP_HOME, HOME  toolchain location (defaults derived from
#                  the passwd entry, so a systemd scope with an empty
#                  environment works the same as a login shell).
#
# Results: $GATE_OUT/summary.txt has one `NN name=status` line per row and a
# final `DONE failed=N` line; $GATE_OUT/NN-name.out holds each row's output.
# Every cargo call goes through tools/gate/bin/cargo (per-command systemd
# memory scope). Exit status is 0 only when every selected row passes.
set -u
HERE=$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)
NAME=${1:?usage: run_gate.sh <name> [row ...] | --list}
shift
GATE_DIR=$(cd "${GATE_DIR:-$HERE/../..}" && pwd) || exit 2
cd "$GATE_DIR" || exit 2

# Toolchain: never rely on a login-shell PATH.
export HOME=${HOME:-$(getent passwd "$(id -u)" | cut -d: -f6)}
export CARGO_HOME=${CARGO_HOME:-$HOME/.cargo}
export RUSTUP_HOME=${RUSTUP_HOME:-$HOME/.rustup}
export PATH="$GATE_DIR/tools/gate/bin:$CARGO_HOME/bin:/usr/local/bin:/usr/bin:/bin"
unset LITCHI_CAPPED RUSTC

export CARGO_BUILD_JOBS=${CARGO_BUILD_JOBS:-8} RUST_TEST_THREADS=${RUST_TEST_THREADS:-8}
export CARGO_TARGET_DIR=$GATE_DIR/target CARGO_INCREMENTAL=0
export CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 CARGO_TERM_COLOR=never
[ "${GATE_OFFLINE:-}" = 1 ] && export CARGO_NET_OFFLINE=true
export TMPDIR=${GATE_TMPDIR:-/tmp/litchi-gate-$(id -u)/$NAME}

S=tools/gate
CHUNKED="python3 $S/chunked_cargo_tests.py"
FEATURES=$(python3 -c 'import tomllib;f=tomllib.load(open("crates/litchi-ops/Cargo.toml","rb"))["features"];print(",".join(k for k in f if k not in ("default","encryption")))')
SUBS=("" xlsx pptx docx xls xlsb doc ppt rtf odt ods odp oth odm odg odc odi odb odf-formula pages keynote numbers)
LEAF="-p litchi-docx -p litchi-pptx -p litchi-xlsx -p litchi-xlsb -p litchi-doc -p litchi-xls -p litchi-ppt -p litchi-rtf -p litchi-odt -p litchi-odp -p litchi-opc -p xml-minifier -p litchi-drawingml -p litchi-spreadsheet-drawing -p litchi-keynote -p litchi-pages -p litchi-numbers -p litchi-numbers-wire -p litchi-iwa-text-wire"
TOOL_MODULES="tools.test_check_api_glossary tools.test_api_debt_inventory tools.test_validate_api_disposition tools.test_final_api_inventory tools.test_raw_api_coverage tools.test_prepare_raw_api_census tools.test_w8_file_provenance tools.test_w8_original_baseline tools.test_w8_projection_policy tools.test_w8_quiet_abba tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references tools.test_check_crate_boundaries tools.test_non_iwork_gate tools.test_unified_bench_matrix tools.test_check_iwork_leaf_public_api tools.test_native_qualification_runner tools.test_legacy_native_qualification tools.test_prepare_xlsx_raw_inputs tools.test_xlsx_raw_native_qualification tools.test_prepare_ole_owner_inputs tools.test_ole_owner_paired_runner tools.test_validate_ole_owner_run tools.test_ppt_lazy_build_receipt_adapter tools.test_ppt_lazy_cold_first_commit tools.test_ppt_lazy_native_qualification tools.test_prepare_ppt_lazy_native_inputs tools.test_drawingml_compile_driver tools.test_drawingml_raw_qualification tools.test_drawingml_raw_policy_v2 tools.test_signed_native_policy_v2 tools.test_current_native_policy_v2 docs/unified-ops/results/0362ul29/tooling/test_safe_chunked_tests.py docs/unified-ops/results/0362ul29/tooling/test_run_safe_gate_ppt.py tools.test_ink_safe_runner docs/unified-ops/results/0362ul24/test_prepare_signed_native_consumer.py tools.test_opc_read_ahead_timing_guard tools.test_xls_f3_qualification_abba docs/unified-ops/results/0362ul24/test_validate_signed_native.py tools.test_gate_runner"
# Every crates/*/fuzz workspace has a tracked Cargo.lock; all are checked --locked.
FUZZ='set -e; export CARGO_TARGET_DIR="$CARGO_TARGET_DIR/fuzz"
for m in crates/*/fuzz/Cargo.toml; do echo "+ check $m"; cargo check -q --locked --manifest-path "$m" --bins; done
cargo check -q --locked --manifest-path crates/litchi-ops/fuzz/Cargo.toml --bins --features apply
cargo run -q --locked --release --manifest-path crates/litchi-ops/fuzz/Cargo.toml --bin smoke_batch_payload
cargo run -q --locked --release --manifest-path crates/litchi-ops/fuzz/Cargo.toml --features apply --bin smoke_batch_payload'

ROWS=()   # "name<TAB>command" (command is run with bash -c)
row() { ROWS+=("$1"$'\t'"$2"); }
row fmt "cargo fmt --all --check"
row clippy-all-features "cargo clippy -q ${GATE_EXTRA:-} -p litchi-cli -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings"
row clippy-ops-no-encryption "cargo clippy -q -p litchi-ops --features '$FEATURES' --all-targets -- -D warnings"
for f in "${SUBS[@]}"; do
    row "clippy-ops-${f:-none}" "cargo clippy -q -p litchi-ops --no-default-features ${f:+--features $f} --all-targets -- -D warnings"
done
row tests-ops "$CHUNKED -p litchi-ops --all-features"
row tests-facade "$CHUNKED -p litchi --all-features"
row tests-core "$CHUNKED -p litchi-core --all-features"
row clippy-leaves "cargo clippy -q $LEAF --all-targets --all-features -- -D warnings"
row clippy-ods "cargo clippy -q -p litchi-ods --all-targets --all-features -- -D warnings"
row tests-leaves "$CHUNKED $LEAF --all-features"
row tests-ods-full "$CHUNKED -p litchi-ods --all-features"
row clippy-workspace "cargo clippy -q --workspace --all-targets --all-features -- -D warnings"
row tests-workspace-excluding-ods-ops "$CHUNKED --workspace --exclude litchi-ods --exclude litchi-ops --all-features"
row wsdoc "CARGO_BUILD_JOBS=1 RUSTDOCFLAGS=-Dwarnings cargo doc -q --workspace --no-deps --all-features"
row ppt-alone "cargo test -q -p litchi-ppt --all-features --lib"
row rustdoc-core-ops-facade "RUSTDOCFLAGS=-Dwarnings cargo doc -q --no-deps --all-features -p litchi-core -p litchi-ops -p litchi"
row cli-tests "$CHUNKED -p litchi-cli --all-features"
row python-bindings "bash crates/pyo3-litchi/tests/run_tests.sh"
row feature-matrix "python3 tools/check_feature_matrix.py"
row feature-matrix-clippy "python3 tools/check_feature_matrix.py --only facade,facade-aggregate --all-targets --clippy"
row fuzz "$FUZZ"
row goldens "bash $S/check_goldens.sh"
row golden-clean 'test -z "$(git status --short | grep -v "^[AMDR] ")"'
if [ -n "${HARNESS_TESTS:-}" ]; then
    row perf-harness-check "cd tools/perf-baseline && cargo test --locked --lib --tests --no-fail-fast"
else
    row perf-harness-check "cd tools/perf-baseline && cargo check --locked --all-targets"
fi
row native-resave-check "cd tools/native-resave && cargo check --locked --all-targets"
row api-glossary 'RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --target-dir "$CARGO_TARGET_DIR"'
row iwork-leaf-api "python3 -m unittest tools.test_check_iwork_leaf_public_api"
row tooling-unit "python3 -m unittest $TOOL_MODULES"
row crate-boundaries "python3 tools/check_crate_boundaries.py"
row non-iwork-gate "python3 tools/non_iwork_gate.py verify"
row bench-matrix "python3 tools/unified_bench_matrix.py --check"
row corpus-dispositions "python3 tools/officecli_capability_delta.py --check && python3 tools/officecli_corpus.py --check"

if [ "${#ROWS[@]}" -ne 53 ]; then echo "gate defines ${#ROWS[@]} rows, expected 53" >&2; exit 2; fi
if [ "$NAME" = --list ]; then
    for i in "${!ROWS[@]}"; do printf '%02d %s\n' "$i" "${ROWS[$i]%%$'\t'*}"; done
    exit 0
fi

selected() {  # selected <index> <name>
    [ "$#" -gt 0 ] || return 1
    [ "${#SELECT[@]}" -eq 0 ] && return 0
    local s
    for s in "${SELECT[@]}"; do
        if [ "$s" = "$2" ] || { [[ "$s" =~ ^[0-9]+$ ]] && [ "$((10#$s))" -eq "$1" ]; }; then return 0; fi
    done
    return 1
}
SELECT=("$@")
G=${GATE_OUT:-$GATE_DIR/target/gate-results/$NAME}
mkdir -p "$TMPDIR" "$G" "$CARGO_TARGET_DIR"

{
    echo "gate=$NAME dir=$GATE_DIR commit=$(git rev-parse HEAD) dirty=$(git status --porcelain | wc -l)"
    echo "started=$(date -u +%FT%TZ)"
    echo "PATH=$PATH"
    echo "cargo=$(cargo -V 2>&1)"
    rustc -vV 2>&1
} > "$G/environment.txt"
if ! rustc -V >/dev/null 2>&1 || ! cargo -V >/dev/null 2>&1; then
    echo "toolchain not found under $CARGO_HOME/bin; see $G/environment.txt" >&2
    exit 2
fi

# Root Cargo.lock: this library workspace does not track it (ADR 0028), but
# several rows (feature matrix, --locked legs) need one. A checkout without a
# lock gets GATE_LOCK_SEED (e.g. the trunk worktree's lock) or, without a
# seed, a freshly resolved one; either way `cargo metadata` then reconciles it
# with the manifests, and the identity is recorded.
if [ ! -f Cargo.lock ]; then
    if [ -n "${GATE_LOCK_SEED:-}" ]; then
        cp "$GATE_LOCK_SEED" Cargo.lock || exit 2
        echo "lock=seeded from $GATE_LOCK_SEED" >> "$G/environment.txt"
    else
        cargo generate-lockfile -q >> "$G/environment.txt" 2>&1 || exit 2
        echo "lock=generated" >> "$G/environment.txt"
    fi
fi
if ! cargo metadata -q --format-version 1 > /dev/null 2>> "$G/environment.txt"; then
    echo "cargo metadata failed; see $G/environment.txt" >&2
    exit 2
fi
echo "lock_sha256=$(sha256sum Cargo.lock | cut -d' ' -f1)" >> "$G/environment.txt"

: > "$G/summary.txt"
failed=0
for i in "${!ROWS[@]}"; do
    name=${ROWS[$i]%%$'\t'*}
    command=${ROWS[$i]#*$'\t'}
    selected "$i" "$name" || continue
    out=$(printf '%s/%02d-%s.out' "$G" "$i" "$name")
    if ! python3 "$S/target_headroom.py" "$CARGO_TARGET_DIR" >> "$G/headroom.out" 2>&1; then
        echo "target headroom check failed; see headroom.out" > "$out"
        status=125
    else
        start=$(date +%s)
        bash -c "$command" > "$out" 2>&1 < /dev/null
        status=$?
        echo "[row $i $name: exit $status after $(( $(date +%s) - start ))s]" >> "$out"
    fi
    printf '%02d %s=%s\n' "$i" "$name" "$status" >> "$G/summary.txt"
    [ "$status" -eq 0 ] || failed=$((failed + 1))
done
echo "DONE failed=$failed finished=$(date -u +%FT%TZ)" >> "$G/summary.txt"
[ "$failed" -eq 0 ]
