#!/usr/bin/env bash
set -euo pipefail

script_dir=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
evidence_dir=$(CDPATH= cd -- "$script_dir/.." && pwd)
source_root=${LITCHI_FORM_OWNER_READ_SOURCE_ROOT:-}
fixture_root=${LITCHI_FORM_OWNER_READ_FIXTURE_ROOT:-}
if [[ -z "$source_root" || -z "$fixture_root" ]]; then
    printf '%s\n' 'set LITCHI_FORM_OWNER_READ_SOURCE_ROOT and LITCHI_FORM_OWNER_READ_FIXTURE_ROOT' >&2
    exit 2
fi
source_root=$(CDPATH= cd -- "$source_root" && pwd)
fixture_root=$(CDPATH= cd -- "$fixture_root" && pwd)
gate_file="$evidence_dir/source-root-gate.json"
[[ -f "$gate_file" ]] || { printf 'missing source gate: %s\n' "$gate_file" >&2; exit 2; }
[[ -f "$evidence_dir/source-manifest.sha256" ]] || { printf 'missing source manifest\n' >&2; exit 2; }
[[ -f "$evidence_dir/source-Cargo.lock" ]] || { printf 'missing source lock\n' >&2; exit 2; }
[[ -f "$script_dir/Cargo.lock" ]] || { printf 'missing harness lock\n' >&2; exit 2; }
cmp "$evidence_dir/source-Cargo.lock" "$source_root/Cargo.lock" || {
    printf '%s\n' 'source Cargo.lock differs from retained source lock' >&2
    exit 2
}
if [[ -f "$evidence_dir/harness-manifest.sha256" ]]; then
    python3 - "$evidence_dir" "$evidence_dir/harness-manifest.sha256" <<'PY'
import hashlib
import pathlib
import sys

root = pathlib.Path(sys.argv[1])
for line in pathlib.Path(sys.argv[2]).read_text().splitlines():
    expected, relative = line.split("  ", 1)
    actual = hashlib.sha256((root / relative).read_bytes()).hexdigest()
    if actual != expected:
        raise SystemExit(f"harness manifest mismatch: {relative}")
PY
fi

runs_dir="$evidence_dir/runs"
mkdir -p "$runs_dir"
if [[ -n "${LITCHI_FORM_OWNER_READ_OUTPUT_DIR:-}" ]]; then
    output_dir=$LITCHI_FORM_OWNER_READ_OUTPUT_DIR
    mkdir -p "$output_dir"
    output_dir=$(CDPATH= cd -- "$output_dir" && pwd)
    if find "$output_dir" -mindepth 1 -print -quit | grep -q .; then
        printf 'refusing to overwrite non-empty output directory: %s\n' "$output_dir" >&2
        exit 2
    fi
else
    output_dir=$(mktemp -d "$runs_dir/replay-$(date -u +%Y%m%dT%H%M%SZ).XXXXXX")
fi

target_dir=$(mktemp -d "${TMPDIR:-/var/tmp}/litchi-form-owner-read-target.XXXXXX")
work_dir=$(mktemp -d "${TMPDIR:-/var/tmp}/litchi-form-owner-read-harness.XXXXXX")
synthetic_dir=$(mktemp -d "${TMPDIR:-/var/tmp}/litchi-form-owner-read-fixtures.XXXXXX")
export CARGO_TARGET_DIR="$target_dir"
cleanup() {
    rm -rf "$target_dir" "$work_dir" "$synthetic_dir"
}
trap cleanup EXIT

expected_commit=$(python3 -c 'import json,sys; print(json.load(open(sys.argv[1]))["commit"])' "$gate_file")
actual_commit=$(git -C "$source_root" rev-parse HEAD)
[[ "$actual_commit" == "$expected_commit" ]] || {
    printf 'source commit mismatch: %s != %s\n' "$actual_commit" "$expected_commit" >&2
    exit 2
}
git -C "$source_root" diff --quiet -- . ':!docs/report/spec-gap-validation-evidence/xlsx-form-control-owner-read-performance' || {
    printf '%s\n' 'source checkout has tracked edits outside the evidence directory' >&2
    exit 2
}
python3 - "$source_root" "$evidence_dir/source-manifest.sha256" "$output_dir/source-state-before.json" <<'PY'
import hashlib
import json
import pathlib
import subprocess
import sys

root = pathlib.Path(sys.argv[1])
manifest = pathlib.Path(sys.argv[2])
output = pathlib.Path(sys.argv[3])
errors = []
count = 0
for line in manifest.read_text().splitlines():
    expected, relative = line.split("  ", 1)
    path = root / relative
    if not path.is_file():
        errors.append(f"missing {relative}")
        continue
    digest = hashlib.sha256(path.read_bytes()).hexdigest()
    count += 1
    if digest != expected:
        errors.append(f"hash mismatch {relative}: {digest} != {expected}")
result = {
    "status": "pass" if not errors else "fail",
    "commit": subprocess.check_output(["git", "-C", str(root), "rev-parse", "HEAD"], text=True).strip(),
    "manifest_entries": count,
    "manifest_sha256": hashlib.sha256(manifest.read_bytes()).hexdigest(),
    "errors": errors,
}
output.write_text(json.dumps(result, indent=2) + "\n")
if errors:
    raise SystemExit("source manifest verification failed")
PY

mkdir -p "$work_dir/package/src" "$work_dir/source"
cp "$script_dir/Cargo.toml" "$work_dir/package/Cargo.toml"
cp "$script_dir/src/main.rs" "$work_dir/package/src/main.rs"
cp "$script_dir/Cargo.lock" "$work_dir/package/Cargo.lock"
cp "$source_root/Cargo.toml" "$work_dir/source/Cargo.toml"
cp "$source_root/Cargo.lock" "$work_dir/source/Cargo.lock"
ln -s "$source_root/crates" "$work_dir/source/crates"
if [[ -d "$source_root/.cargo" ]]; then
    ln -s "$source_root/.cargo" "$work_dir/source/.cargo"
    ln -s "$source_root/.cargo" "$work_dir/package/.cargo"
fi
for config in clippy.toml rust-toolchain.toml; do
    if [[ -f "$source_root/$config" ]]; then
        ln -s "$source_root/$config" "$work_dir/package/$config"
    fi
done

(cd "$work_dir/package" && cargo build --release --locked --offline --manifest-path Cargo.toml)
binary="$target_dir/release/form-owner-perf-harness"
sha256sum "$binary" > "$output_dir/harness-binary-before.sha256"

python3 "$script_dir/generate_many_controls.py" \
    "$fixture_root/button-form-control.xlsx" "$synthetic_dir/many-controls-16.xlsx" 16
python3 "$script_dir/generate_cases.py" \
    "$fixture_root/tdf161365.xlsx" "$synthetic_dir/diagnostic-mirror.xlsx" DIAGNOSTIC
python3 "$script_dir/generate_cases.py" \
    "$fixture_root/checkbox-form-control.xlsx" "$synthetic_dir/foreign-wrapper-refusal.xlsx" REFUSAL

python3 - "$fixture_root" "$synthetic_dir" "$output_dir/fixture-hashes.sha256" "$output_dir/fixture-member-hashes.sha256" <<'PY'
import hashlib
import pathlib
import sys
import zipfile

fixture_root = pathlib.Path(sys.argv[1])
synthetic_root = pathlib.Path(sys.argv[2])
fixture_paths = [fixture_root / "button-form-control.xlsx", fixture_root / "tdf60673.xlsx", fixture_root / "tdf161365.xlsx"]
fixture_paths.extend(sorted(synthetic_root.glob("*.xlsx")))
lines = []
member_lines = []
for path in fixture_paths:
    lines.append(f"{hashlib.sha256(path.read_bytes()).hexdigest()}  {path.name}")
    with zipfile.ZipFile(path, "r") as archive:
        for info in archive.infolist():
            member_lines.append(f"{hashlib.sha256(archive.read(info.filename)).hexdigest()}  {path.name}!{info.filename}")
pathlib.Path(sys.argv[3]).write_text("\n".join(lines) + "\n")
pathlib.Path(sys.argv[4]).write_text("\n".join(member_lines) + "\n")
PY

raw_receipts="$output_dir/raw-receipts.jsonl"
: > "$raw_receipts"
run_fixture() {
    local label=$1
    local path=$2
    local count=$3
    local diagnostics=$4
    "$binary" --fixture "$path" --label "$label" --expected-count "$count" \
        --expected-diagnostics "$diagnostics" --iterations 7 >> "$raw_receipts"
}
run_refusal() {
    local label=$1
    local path=$2
    "$binary" --fixture "$path" --label "$label" --expect-refusal --iterations 7 >> "$raw_receipts"
}

run_fixture button-small "$fixture_root/button-form-control.xlsx" 1 0
run_fixture tdf60673-medium "$fixture_root/tdf60673.xlsx" 2 0
run_fixture many-controls-16 "$synthetic_dir/many-controls-16.xlsx" 16 0
run_fixture diagnostic-mirror "$synthetic_dir/diagnostic-mirror.xlsx" 2 1
run_refusal foreign-wrapper-refusal "$synthetic_dir/foreign-wrapper-refusal.xlsx"

sha256sum "$binary" > "$output_dir/harness-binary-after.sha256"
cmp "$output_dir/harness-binary-before.sha256" "$output_dir/harness-binary-after.sha256"
cp "$output_dir/source-state-before.json" "$output_dir/source-state-after.json"

python3 - "$raw_receipts" "$output_dir/receipt-index.json" <<'PY'
import json
import pathlib
import sys

records = [json.loads(line) for line in pathlib.Path(sys.argv[1]).read_text().splitlines()]
receipts = [record for record in records if record["record"] == "receipt"]
correctness = [record for record in records if record["record"] in {"correctness", "refusal_correctness"}]
required = {"elapsed_ns", "requested_event_bytes", "live_bytes_delta", "peak_live_bytes", "rss_before_bytes", "rss_after_bytes"}
if receipts and required - set(receipts[0]):
    raise SystemExit(f"missing receipt fields: {sorted(required - set(receipts[0]))}")
summary = {
    "status": "pass",
    "receipt_count": len(receipts),
    "correctness_record_count": len(correctness),
    "iterations_per_case": 7,
    "accepted_cases": ["button-small", "tdf60673-medium", "many-controls-16", "diagnostic-mirror"],
    "refusal_cases": ["foreign-wrapper-refusal"],
    "lanes": sorted({record["lane"] for record in receipts}),
    "allocator_metric_definition": {
        "requested_event_bytes": "cumulative successful alloc/realloc requested sizes during the phase",
        "live_bytes_delta": "phase-end live-byte change from the pre-phase allocator baseline",
        "peak_live_bytes": "maximum live-byte increase above that baseline during the phase",
        "rss": "separate process VmRSS snapshots where available",
    },
    "candidate_only": True,
    "baseline_claim": False,
    "speedup_claim": False,
}
pathlib.Path(sys.argv[2]).write_text(json.dumps(summary, indent=2) + "\n")
PY

python3 - "$raw_receipts" "$output_dir/sanity-validation.json" <<'PY'
import json
import pathlib
import sys

records = [json.loads(line) for line in pathlib.Path(sys.argv[1]).read_text().splitlines()]
receipts = [record for record in records if record["record"] == "receipt"]
accepted = [record for record in records if record["record"] == "correctness"]
refusals = [record for record in records if record["record"] == "refusal_correctness"]
required = {
    "eager_owner_projection", "source_owner_projection",
    "eager_package_open", "source_package_open",
    "eager_warm_single_query", "source_warm_single_query",
    "eager_warm_many_query", "source_warm_many_query",
    "eager_noop_clone", "source_noop_clone",
    "eager_refusal", "source_refusal",
}
actual = {record["lane"] for record in receipts}
if len(accepted) != 4 or len(refusals) != 1:
    raise SystemExit(f"unexpected correctness records: accepted={len(accepted)} refusals={len(refusals)}")
if len(receipts) != 294:
    raise SystemExit(f"unexpected receipt count: {len(receipts)}")
if actual != required:
    raise SystemExit(f"lane set mismatch: {sorted(actual)}")
for record in receipts:
    if any(key not in record for key in ("requested_event_bytes", "live_bytes_delta", "peak_live_bytes")):
        raise SystemExit("allocator live fields are incomplete")
for record in receipts:
    if record["lane"] in {
        "eager_warm_single_query", "source_warm_single_query",
        "eager_warm_many_query", "source_warm_many_query",
        "eager_noop_clone", "source_noop_clone",
    } and (
        record["alloc_calls"] or record["dealloc_calls"] or record["realloc_calls"]
        or record["requested_event_bytes"] or record["live_bytes_delta"]
        or record["peak_live_bytes"] or record["source_read_calls"]
    ):
        raise SystemExit(f"warm query/clone allocation or source-read invariant failed: {record}")
diagnostic = [record for record in accepted if record["fixture"] == "diagnostic-mirror"][0]
if diagnostic["diagnostics"] < 1:
    raise SystemExit("diagnostic fixture did not retain an admitted diagnostic")
if len({record["eager_error_kind"] for record in refusals}) != 1:
    raise SystemExit(f"unexpected eager refusal kinds: {refusals}")
if len({record["source_error_kind"] for record in refusals}) != 1:
    raise SystemExit(f"unexpected source refusal kinds: {refusals}")
if refusals[0]["eager_error_kind"] != refusals[0]["source_error_kind"]:
    raise SystemExit(f"eager/source refusal error class diverged: {refusals}")
summary = {
    "status": "pass",
    "candidate_only": True,
    "accepted_correctness_records": len(accepted),
    "refusal_correctness_records": len(refusals),
    "receipt_records": len(receipts),
    "iterations_per_case": 7,
    "warm_query_clone_zero_allocator_and_source_reads": True,
    "diagnostics_measured_separately": True,
    "unsupported_graph_refusal_measured_separately": True,
    "baseline_claim": False,
    "speedup_claim": False,
}
pathlib.Path(sys.argv[2]).write_text(json.dumps(summary, indent=2) + "\n")
PY

{
    printf 'source_root=%s\n' "$source_root"
    printf 'commit=%s\n' "$expected_commit"
    printf 'output_dir=%s\n' "$output_dir"
    printf 'target_dir_isolated=true\n'
    printf 'harness_lock_sha256=%s\n' "$(sha256sum "$script_dir/Cargo.lock" | cut -d' ' -f1)"
    printf 'source_lock_sha256=%s\n' "$(sha256sum "$evidence_dir/source-Cargo.lock" | cut -d' ' -f1)"
    printf 'harness_binary_sha256=%s\n' "$(cut -d' ' -f1 "$output_dir/harness-binary-after.sha256")"
    printf 'raw_receipts_sha256=%s\n' "$(sha256sum "$raw_receipts" | cut -d' ' -f1)"
    printf 'fixture_hashes_sha256=%s\n' "$(sha256sum "$output_dir/fixture-hashes.sha256" | cut -d' ' -f1)"
    printf '\n'
    rustc -Vv
    cargo -V
    uname -a
    if command -v lscpu >/dev/null 2>&1; then lscpu; fi
} > "$output_dir/build-provenance.txt"

printf 'candidate run output: %s\n' "$output_dir" >&2
