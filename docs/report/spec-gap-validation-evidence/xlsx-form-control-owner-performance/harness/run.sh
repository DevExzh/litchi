#!/usr/bin/env bash
set -euo pipefail

script_dir=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
evidence_dir=$(CDPATH= cd -- "$script_dir/.." && pwd)
source_root=${LITCHI_FORM_OWNER_SOURCE_ROOT:-}
fixture_root=${LITCHI_FORM_OWNER_FIXTURE_ROOT:-}
if [[ -z "$source_root" || -z "$fixture_root" ]]; then
    printf '%s\n' 'set LITCHI_FORM_OWNER_SOURCE_ROOT and LITCHI_FORM_OWNER_FIXTURE_ROOT' >&2
    exit 2
fi

source_root=$(CDPATH= cd -- "$source_root" && pwd)
fixture_root=$(CDPATH= cd -- "$fixture_root" && pwd)
gate_file="$evidence_dir/source-root-gate.json"
[[ -f "$gate_file" ]] || { printf 'missing frozen source gate: %s\n' "$gate_file" >&2; exit 2; }
[[ -f "$evidence_dir/source-Cargo.lock" ]] || {
    printf 'missing retained source lock: %s\n' "$evidence_dir/source-Cargo.lock" >&2
    exit 2
}
[[ -f "$script_dir/Cargo.lock" ]] || {
    printf 'missing retained harness lock: %s\n' "$script_dir/Cargo.lock" >&2
    exit 2
}

runs_dir="$evidence_dir/runs"
mkdir -p "$runs_dir"
if [[ -n "${LITCHI_FORM_OWNER_OUTPUT_DIR:-}" ]]; then
    output_dir=$LITCHI_FORM_OWNER_OUTPUT_DIR
    mkdir -p "$output_dir"
    output_dir=$(CDPATH= cd -- "$output_dir" && pwd)
    if find "$output_dir" -mindepth 1 -print -quit | grep -q .; then
        printf 'refusing to overwrite non-empty output directory: %s\n' "$output_dir" >&2
        exit 2
    fi
else
    output_dir=$(mktemp -d "$runs_dir/replay-$(date -u +%Y%m%dT%H%M%SZ).XXXXXX")
fi

target_dir=$(mktemp -d "${TMPDIR:-/tmp}/litchi-form-owner-target.XXXXXX")
work_dir=$(mktemp -d "${TMPDIR:-/tmp}/litchi-form-owner-harness.XXXXXX")
synthetic_dir=$(mktemp -d "${TMPDIR:-/tmp}/litchi-form-owner-fixtures.XXXXXX")
export CARGO_TARGET_DIR="$target_dir"
cleanup() {
    rm -rf "$target_dir" "$work_dir" "$synthetic_dir"
}
trap cleanup EXIT

verify_frozen_state() {
    local phase=$1
    python3 - "$gate_file" "$source_root" "$script_dir" "$evidence_dir" \
        "$output_dir/source-state-${phase}.json" \
        "$output_dir/source-manifest-${phase}.sha256" <<'PY'
import hashlib
import json
import os
import pathlib
import subprocess
import sys

(
    gate_path,
    source_root_arg,
    harness_root_arg,
    evidence_root_arg,
    output_arg,
    manifest_output_arg,
) = sys.argv[1:]
gate = json.loads(pathlib.Path(gate_path).read_text())
source_root = pathlib.Path(source_root_arg)
harness_root = pathlib.Path(harness_root_arg)
evidence_root = pathlib.Path(evidence_root_arg)
errors = []

def digest(path):
    hasher = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            hasher.update(chunk)
    return hasher.hexdigest()

source_hashes = {}
for relative, expected in gate["sha256"].items():
    path = source_root / relative
    if not path.is_file():
        errors.append(f"missing source file: {relative}")
        continue
    actual = digest(path)
    source_hashes[relative] = actual
    if actual != expected:
        errors.append(f"source hash mismatch: {relative}: {actual} != {expected}")

def expected_hash_file(name):
    return (evidence_root / name).read_text().split()[0]

source_lock_expected = expected_hash_file("source-Cargo.lock.sha256")
source_lock_actual = digest(source_root / "Cargo.lock")
if source_lock_actual != source_lock_expected:
    errors.append(f"source Cargo.lock hash mismatch: {source_lock_actual} != {source_lock_expected}")
retained_source_lock = digest(evidence_root / "source-Cargo.lock")
if retained_source_lock != source_lock_expected:
    errors.append("retained source-Cargo.lock does not match its recorded hash")

harness_files = {}
for relative in ("Cargo.toml", "src/main.rs", "run.sh", "Cargo.lock"):
    path = harness_root / relative
    if not path.is_file():
        errors.append(f"missing harness file: {relative}")
        continue
    harness_files[relative] = digest(path)

def verify_manifest(name, root, manifest):
    values = []
    path = evidence_root / manifest
    if not path.is_file():
        errors.append(f"missing pinned manifest: {manifest}")
        return values
    for line in path.read_text().splitlines():
        expected, relative = line.split("  ", 1)
        candidate = root / relative
        if not candidate.is_file():
            errors.append(f"missing manifest file: {relative}")
            continue
        actual = digest(candidate)
        values.append(f"{actual}  {relative}")
        if actual != expected:
            errors.append(f"manifest hash mismatch: {relative}: {actual} != {expected}")
    return values

source_manifest_lines = verify_manifest(
    "source-manifest.sha256", source_root, "source-manifest.sha256"
)
source_config_lines = verify_manifest(
    "source-build-config.sha256", source_root, "source-build-config.sha256"
)
harness_manifest_lines = verify_manifest(
    "harness-manifest.sha256", harness_root, "harness-manifest.sha256"
)
pathlib.Path(manifest_output_arg).write_text("\n".join(source_manifest_lines) + "\n")

def command_receipt(command):
    try:
        return subprocess.check_output(
            command, cwd=source_root, stderr=subprocess.STDOUT, text=True
        ).strip()
    except (OSError, subprocess.CalledProcessError) as error:
        errors.append(f"cannot capture {' '.join(command)}: {error}")
        return None

compiler = {
    "rustc_vv": command_receipt(["rustc", "-Vv"]),
    "cargo_v": command_receipt(["cargo", "-V"]),
    "rustup_active_toolchain": command_receipt(["rustup", "show", "active-toolchain"]),
}
environment = {
    key: os.environ.get(key)
    for key in (
        "TMPDIR",
        "CARGO_TARGET_DIR",
        "CARGO_HOME",
        "RUSTFLAGS",
        "RUSTC_WRAPPER",
        "RUSTUP_TOOLCHAIN",
        "LITCHI_FORM_OWNER_SOURCE_ROOT",
        "LITCHI_FORM_OWNER_FIXTURE_ROOT",
    )
}

artifact_hashes = {}
for name in (
    "source-bundle.tar",
    "source-delta.patch",
    "source-Cargo.lock",
    "harness/Cargo.lock",
    "source-manifest.sha256",
    "source-build-config.sha256",
    "harness-manifest.sha256",
):
    path = evidence_root / name
    if not path.is_file():
        errors.append(f"missing provenance artifact: {name}")
        continue
    sidecar = evidence_root / f"{name.replace('/', '-')}.sha256"
    if name == "source-Cargo.lock":
        sidecar = evidence_root / "source-Cargo.lock.sha256"
    if name == "harness/Cargo.lock":
        sidecar = evidence_root / "harness-Cargo.lock.sha256"
    actual = digest(path)
    artifact_hashes[name] = actual
    if sidecar.is_file() and actual != sidecar.read_text().split()[0]:
        errors.append(f"provenance hash mismatch: {name}")

git_metadata = None
git_dir = source_root / ".git"
if git_dir.exists():
    try:
        head = subprocess.check_output(
            ["git", "-C", str(source_root), "rev-parse", "HEAD"], text=True
        ).strip()
        status = subprocess.check_output(
            ["git", "-C", str(source_root), "status", "--porcelain", "--untracked-files=all"],
            text=True,
        ).splitlines()
        status_paths = {line[3:] for line in status if len(line) >= 4}
        expected_paths = set(gate["files"])
        if head != gate["base"]:
            errors.append(f"source HEAD mismatch: {head} != {gate['base']}")
        if status_paths != expected_paths:
            errors.append(f"source dirty path set mismatch: {sorted(status_paths)}")
        git_metadata = {"head": head, "status": status}
    except (OSError, subprocess.CalledProcessError) as error:
        errors.append(f"cannot inspect source git metadata: {error}")

result = {
    "status": "pass" if not errors else "fail",
    "source_root": str(source_root),
    "base": gate["base"],
    "source_hashes": source_hashes,
    "source_Cargo_lock_sha256": source_lock_actual,
    "harness_hashes": harness_files,
    "source_manifest_count": len(source_manifest_lines),
    "source_manifest_sha256": digest(pathlib.Path(manifest_output_arg)),
    "source_build_config_manifest_count": len(source_config_lines),
    "harness_manifest_count": len(harness_manifest_lines),
    "provenance_hashes": artifact_hashes,
    "compiler": compiler,
    "environment": environment,
    "git": git_metadata,
    "errors": errors,
}
pathlib.Path(output_arg).write_text(json.dumps(result, indent=2) + "\n")
if errors:
    raise SystemExit("frozen state verification failed; see " + output_arg)
PY
}

verify_frozen_state before

mkdir -p "$work_dir/package/src" "$work_dir/source"
cp "$script_dir/Cargo.toml" "$work_dir/package/Cargo.toml"
cp "$script_dir/src/main.rs" "$work_dir/package/src/main.rs"
cp "$script_dir/Cargo.lock" "$work_dir/package/Cargo.lock"
cp "$source_root/Cargo.toml" "$work_dir/source/Cargo.toml"
cp "$source_root/Cargo.lock" "$work_dir/source/Cargo.lock"
ln -s "$source_root/crates" "$work_dir/source/crates"
if [[ -d "$source_root/.cargo" ]]; then
    ln -s "$source_root/.cargo" "$work_dir/source/.cargo"
fi
if [[ -d "$source_root/.cargo" ]]; then
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

python3 - "$fixture_root/button-form-control.xlsx" "$synthetic_dir/unrelated-opaque-1MiB.xlsx" <<'PY'
import hashlib
import sys
import zipfile

source, destination = sys.argv[1:]
payload = b"".join(hashlib.sha256(i.to_bytes(8, "little")).digest() for i in range(32768))
fixed_date = (1980, 1, 1, 0, 0, 0)

def fixed_info(info):
    clone = zipfile.ZipInfo(info.filename, fixed_date)
    clone.compress_type = info.compress_type
    clone.comment = info.comment
    clone.extra = info.extra
    clone.create_system = info.create_system
    clone.create_version = info.create_version
    clone.extract_version = info.extract_version
    clone.flag_bits = info.flag_bits
    clone.volume = info.volume
    clone.internal_attr = info.internal_attr
    clone.external_attr = info.external_attr
    return clone

with zipfile.ZipFile(source, "r") as source_zip, zipfile.ZipFile(
    destination, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9
) as destination_zip:
    for info in source_zip.infolist():
        destination_zip.writestr(
            fixed_info(info),
            source_zip.read(info.filename),
            compress_type=info.compress_type,
            compresslevel=9,
        )
    unrelated = zipfile.ZipInfo("xl/opaque/unrelated.bin", fixed_date)
    unrelated.compress_type = zipfile.ZIP_DEFLATED
    unrelated.create_system = 0
    unrelated.external_attr = 0
    destination_zip.writestr(unrelated, payload, compress_type=zipfile.ZIP_DEFLATED, compresslevel=9)
PY

python3 - "$fixture_root" "$synthetic_dir/unrelated-opaque-1MiB.xlsx" "$output_dir/fixture-hashes.sha256" <<'PY'
import hashlib
import pathlib
import sys

fixture_root = pathlib.Path(sys.argv[1])
synthetic = pathlib.Path(sys.argv[2])
paths = sorted(fixture_root.glob("*.xlsx"))
paths.append(synthetic)
lines = []
for path in paths:
    lines.append(f"{hashlib.sha256(path.read_bytes()).hexdigest()}  {path.name}")
pathlib.Path(sys.argv[3]).write_text("\n".join(lines) + "\n")
PY

python3 - "$fixture_root" "$synthetic_dir/unrelated-opaque-1MiB.xlsx" "$output_dir/fixture-member-hashes.sha256" <<'PY'
import hashlib
import pathlib
import sys
import zipfile

fixture_root = pathlib.Path(sys.argv[1])
synthetic = pathlib.Path(sys.argv[2])
paths = sorted(fixture_root.glob("*.xlsx")) + [synthetic]
lines = []
for path in paths:
    with zipfile.ZipFile(path, "r") as archive:
        for info in archive.infolist():
            digest = hashlib.sha256(archive.read(info.filename)).hexdigest()
            lines.append(f"{digest}  {path.name}!{info.filename}")
pathlib.Path(sys.argv[3]).write_text("\n".join(lines) + "\n")
PY

raw_receipts="$output_dir/raw-receipts.jsonl"
: > "$raw_receipts"
run_fixture() {
    local label=$1
    local path=$2
    local count=$3
    "$binary" --fixture "$path" --label "$label" --expected-count "$count" --iterations 7 \
        >> "$raw_receipts"
}

run_fixture button-form-control "$fixture_root/button-form-control.xlsx" 1
run_fixture checkbox-form-control "$fixture_root/checkbox-form-control.xlsx" 1
run_fixture singlecontrol "$fixture_root/singlecontrol.xlsx" 1
run_fixture tdf120301_xmlSpaceParsing "$fixture_root/tdf120301_xmlSpaceParsing.xlsx" 2
run_fixture tdf134769 "$fixture_root/tdf134769.xlsx" 1
run_fixture tdf161365 "$fixture_root/tdf161365.xlsx" 2
run_fixture tdf60673 "$fixture_root/tdf60673.xlsx" 2
run_fixture unrelated-opaque-1MiB "$synthetic_dir/unrelated-opaque-1MiB.xlsx" 1

sha256sum "$binary" > "$output_dir/harness-binary-after.sha256"
cmp "$output_dir/harness-binary-before.sha256" "$output_dir/harness-binary-after.sha256"
verify_frozen_state after
cmp "$output_dir/source-state-before.json" "$output_dir/source-state-after.json"

python3 - "$raw_receipts" "$output_dir/receipt-index.json" <<'PY'
import json
import pathlib
import sys

receipts = []
correctness = []
for line in pathlib.Path(sys.argv[1]).read_text().splitlines():
    record = json.loads(line)
    if record["record"] == "receipt":
        receipts.append(record)
    elif record["record"] == "correctness":
        correctness.append(record)
required = {"requested_event_bytes", "live_bytes_delta", "peak_live_bytes"}
missing = sorted(required - set(receipts[0])) if receipts else sorted(required)
if missing:
    raise SystemExit(f"missing allocator fields: {missing}")
summary = {
    "status": "pass",
    "receipt_count": len(receipts),
    "correctness_record_count": len(correctness),
    "iterations_per_fixture": 7,
    "fixtures": sorted({record["fixture"] for record in receipts}),
    "lanes": sorted({record["lane"] for record in receipts}),
    "allocator_metric_definition": {
        "requested_event_bytes": "cumulative alloc/realloc requested sizes during the phase",
        "live_bytes_delta": "phase-end live-byte change from the pre-phase global allocator baseline",
        "peak_live_bytes": "maximum live-byte increase above the pre-phase baseline during the phase",
        "rss": "separate process VmRSS snapshots where available",
    },
    "speedup_claim": False,
    "baseline_claim": False,
}
pathlib.Path(sys.argv[2]).write_text(json.dumps(summary, indent=2) + "\n")
PY

python3 - "$raw_receipts" "$output_dir/sanity-validation.json" <<'PY'
import json
import pathlib
import statistics
import sys

records = [json.loads(line) for line in pathlib.Path(sys.argv[1]).read_text().splitlines()]
receipts = [record for record in records if record["record"] == "receipt"]
correctness = [record for record in records if record["record"] == "correctness"]
required_lanes = {
    "eager_cheap_clone", "eager_many_query", "eager_owner_projection", "eager_package_open",
    "eager_single_query", "source_cheap_clone", "source_many_query", "source_owner_projection",
    "source_package_open", "source_single_query",
}
if len(correctness) != 8 or len(receipts) != 560:
    raise SystemExit(f"unexpected record counts: correctness={len(correctness)} receipts={len(receipts)}")
if {record["lane"] for record in receipts} != required_lanes:
    raise SystemExit("required lane set is incomplete")
if any("peak_requested_event_bytes" in record for record in receipts):
    raise SystemExit("new receipts retain the misleading requested-event peak field")
if any(
    record["alloc_calls"] or record["dealloc_calls"] or record["realloc_calls"]
    or record["requested_event_bytes"] or record["live_bytes_delta"] or record["peak_live_bytes"]
    or record["source_read_calls"]
    for record in receipts
    if record["lane"] in {"eager_single_query", "eager_many_query", "eager_cheap_clone", "source_single_query", "source_many_query", "source_cheap_clone"}
):
    raise SystemExit("query/clone lane allocation or source-read invariant failed")

def median(fixture, lane, key):
    values = [
        record[key] for record in receipts
        if record["fixture"] == fixture and record["lane"] == lane
    ]
    return statistics.median(values)

native = "button-form-control"
synthetic = "unrelated-opaque-1MiB"
summary = {
    "status": "pass",
    "candidate_only": True,
    "pre_correction_source": True,
    "correctness_records": len(correctness),
    "receipt_records": len(receipts),
    "iterations_per_fixture": 7,
    "required_lanes": sorted(required_lanes),
    "query_clone_zero_allocation_invariant": True,
    "source_phase_read_counters_present": all("source_read_calls" in record for record in receipts),
    "allocator_live_accounting": True,
    "synthetic_unrelated_member": {
        "median_native_source_package_open_read_bytes": median(native, "source_package_open", "source_read_bytes"),
        "median_synthetic_source_package_open_read_bytes": median(synthetic, "source_package_open", "source_read_bytes"),
        "median_native_source_projection_read_bytes": median(native, "source_owner_projection", "source_read_bytes"),
        "median_synthetic_source_projection_read_bytes": median(synthetic, "source_owner_projection", "source_read_bytes"),
        "median_native_eager_package_open_requested_bytes": median(native, "eager_package_open", "requested_event_bytes"),
        "median_synthetic_eager_package_open_requested_bytes": median(synthetic, "eager_package_open", "requested_event_bytes"),
        "median_native_eager_projection_requested_bytes": median(native, "eager_owner_projection", "requested_event_bytes"),
        "median_synthetic_eager_projection_requested_bytes": median(synthetic, "eager_owner_projection", "requested_event_bytes"),
        "interpretation": "recorded scoped observations only; no baseline or speedup claim",
    },
    "speedup_claim": False,
    "baseline_claim": False,
}
pathlib.Path(sys.argv[2]).write_text(json.dumps(summary, indent=2) + "\n")
PY

{
    printf 'source_root=%s\n' "$source_root"
    printf 'base_commit=%s\n' "$(python3 -c 'import json,sys; print(json.load(open(sys.argv[1]))["base"])' "$gate_file")"
    printf 'output_dir=%s\n' "$output_dir"
    printf 'harness_command=TMPDIR=%q LITCHI_FORM_OWNER_SOURCE_ROOT=%q LITCHI_FORM_OWNER_FIXTURE_ROOT=%q LITCHI_FORM_OWNER_OUTPUT_DIR=%q %q\n' \
        "${TMPDIR:-/tmp}" "$source_root" "$fixture_root" "$output_dir" "$script_dir/run.sh"
    printf 'source_delta_sha256=%s\n' "$(cut -d' ' -f1 "$evidence_dir/source-delta.patch.sha256")"
    printf 'source_bundle_sha256=%s\n' "$(cut -d' ' -f1 "$evidence_dir/source-bundle.tar.sha256")"
    printf 'source_lock_sha256=%s\n' "$(cut -d' ' -f1 "$evidence_dir/source-Cargo.lock.sha256")"
    printf 'harness_lock_sha256=%s\n' "$(cut -d' ' -f1 "$evidence_dir/harness-Cargo.lock.sha256")"
    printf 'harness_binary_before=%s\n' "$(cut -d' ' -f1 "$output_dir/harness-binary-before.sha256")"
    printf 'harness_binary_after=%s\n' "$(cut -d' ' -f1 "$output_dir/harness-binary-after.sha256")"
    printf 'raw_receipts_sha256=%s\n' "$(sha256sum "$raw_receipts" | cut -d' ' -f1)"
    printf 'harness_source_sha256=%s\n' "$(sha256sum "$script_dir/src/main.rs" | cut -d' ' -f1)"
    printf '\n'
    rustc -Vv
    cargo -V
    uname -a
    if command -v lscpu >/dev/null 2>&1; then
        lscpu
    fi
} > "$output_dir/build-provenance.txt"

printf 'candidate run output: %s\n' "$output_dir" >&2
