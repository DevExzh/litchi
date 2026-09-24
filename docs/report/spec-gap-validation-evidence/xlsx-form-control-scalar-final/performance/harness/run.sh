#!/usr/bin/env bash
set -euo pipefail

# This script deliberately refuses to collect a result unless the root agent
# supplies an immutable-source freeze receipt and an explicit capture token.
# A mutable worktree is useful for compile/smoke checks, but it is not final
# performance evidence.
script_dir=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
evidence_dir=$(CDPATH= cd -- "$script_dir/.." && pwd)
source_root=${LITCHI_FORM_SCALAR_SOURCE_ROOT:-}
fixture_root=${LITCHI_FORM_SCALAR_FIXTURE_ROOT:-}
freeze_receipt=${LITCHI_FORM_SCALAR_FREEZE_RECEIPT:-}
capture_token=${LITCHI_FORM_SCALAR_CAPTURE_TOKEN:-}
warmups=${LITCHI_FORM_SCALAR_WARMUPS:-3}
iterations=${LITCHI_FORM_SCALAR_ITERATIONS:-15}
if [[ -z "$source_root" || -z "$fixture_root" || -z "$freeze_receipt" ]]; then
    printf '%s\n' 'set LITCHI_FORM_SCALAR_SOURCE_ROOT, LITCHI_FORM_SCALAR_FIXTURE_ROOT, and LITCHI_FORM_SCALAR_FREEZE_RECEIPT' >&2
    exit 2
fi
if [[ "$capture_token" != "final-source-frozen" ]]; then
    printf '%s\n' 'refusing capture: set LITCHI_FORM_SCALAR_CAPTURE_TOKEN=final-source-frozen after root freeze' >&2
    exit 2
fi
[[ "$warmups" =~ ^[0-9]+$ && "$iterations" =~ ^[1-9][0-9]*$ ]] || {
    printf 'warmups must be a non-negative integer and iterations a positive integer\n' >&2
    exit 2
}
source_root=$(CDPATH= cd -- "$source_root" && pwd)
fixture_root=$(CDPATH= cd -- "$fixture_root" && pwd)
freeze_receipt=$(CDPATH= cd -- "$(dirname -- "$freeze_receipt")" && pwd)/$(basename -- "$freeze_receipt")
[[ -f "$freeze_receipt" ]] || { printf 'missing freeze receipt: %s\n' "$freeze_receipt" >&2; exit 2; }
[[ -f "$source_root/Cargo.toml" ]] || { printf 'missing source Cargo.toml: %s\n' "$source_root" >&2; exit 2; }
[[ -f "$source_root/Cargo.lock" ]] || { printf 'missing source Cargo.lock: %s\n' "$source_root" >&2; exit 2; }
[[ -f "$script_dir/Cargo.lock" ]] || { printf 'missing harness Cargo.lock: %s\n' "$script_dir/Cargo.lock" >&2; exit 2; }

python3 - "$freeze_receipt" "$source_root" <<'PY'
import hashlib
import json
import pathlib
import sys

receipt_path = pathlib.Path(sys.argv[1])
source_root = pathlib.Path(sys.argv[2])
receipt = json.loads(receipt_path.read_text(encoding="utf-8"))
expected_root = receipt.get("source_root")
if expected_root and pathlib.Path(expected_root).resolve() != source_root.resolve():
    raise SystemExit(f"freeze receipt source root mismatch: {expected_root} != {source_root}")
selected = receipt.get("selected_files")
if not isinstance(selected, dict) or not selected:
    raise SystemExit("freeze receipt has no selected_files hash map")
for relative, expected in selected.items():
    candidate = source_root / relative
    if not candidate.is_file():
        raise SystemExit(f"freeze receipt file is missing: {relative}")
    actual = hashlib.sha256(candidate.read_bytes()).hexdigest()
    if actual != expected:
        raise SystemExit(f"freeze receipt hash mismatch for {relative}: {actual} != {expected}")
PY

output_dir=${LITCHI_FORM_SCALAR_OUTPUT_DIR:-$evidence_dir/results/final-$(date -u +%Y%m%dT%H%M%SZ)}
mkdir -p "$output_dir"
output_dir=$(CDPATH= cd -- "$output_dir" && pwd)
if find "$output_dir" -mindepth 1 -print -quit | grep -q .; then
    printf 'refusing to overwrite non-empty output directory: %s\n' "$output_dir" >&2
    exit 2
fi

target_dir=$(mktemp -d "${TMPDIR:-/tmp}/litchi-form-scalar-target.XXXXXX")
work_dir=$(mktemp -d "${TMPDIR:-/tmp}/litchi-form-scalar-harness.XXXXXX")
cleanup() {
    python3 - "$target_dir" "$work_dir" <<'PY'
import os
import pathlib
import shutil
import sys

allowed_parent = pathlib.Path(os.environ.get("TMPDIR", "/tmp")).resolve()
allowed_prefixes = ("litchi-form-scalar-target.", "litchi-form-scalar-harness.")
for raw in sys.argv[1:]:
    candidate = pathlib.Path(raw)
    if candidate.parent.resolve() != allowed_parent:
        raise SystemExit(f"refusing cleanup outside TMPDIR: {candidate}")
    if not candidate.name.startswith(allowed_prefixes):
        raise SystemExit(f"refusing cleanup of unexpected path: {candidate}")
    if candidate.is_symlink():
        raise SystemExit(f"refusing cleanup of symlink: {candidate}")
    if candidate.is_dir():
        shutil.rmtree(candidate)
    elif candidate.exists():
        raise SystemExit(f"refusing cleanup of non-directory: {candidate}")
PY
}
trap cleanup EXIT

manifest() {
    local root=$1
    local output=$2
    python3 - "$root" "$output" <<'PY'
import hashlib
import pathlib
import sys

root = pathlib.Path(sys.argv[1])
destination = pathlib.Path(sys.argv[2])
files = []
for relative in (
    "Cargo.toml",
    "Cargo.lock",
    "rust-toolchain.toml",
    "clippy.toml",
    ".cargo/config.toml",
    ".cargo/config",
):
    candidate = root / relative
    if candidate.is_file():
        files.append(candidate)
for relative_root in (
    "crates/litchi-core",
    "crates/litchi-opc",
    "crates/litchi-ooxml-common",
    "crates/litchi-drawingml",
    "crates/litchi-sheet",
    "crates/litchi-spreadsheet-drawing",
    "crates/litchi-xldm",
    "crates/litchi-xlsx",
    "crates/soapberry-zip",
    "crates/xml-minifier",
    "crates/xml-minifier-macros",
):
    base = root / relative_root
    if base.is_dir():
        files.extend(path for path in base.rglob("*") if path.is_file())
files = sorted(set(files), key=lambda path: path.relative_to(root).as_posix())
with destination.open("w", encoding="utf-8") as stream:
    for path in files:
        digest = hashlib.sha256(path.read_bytes()).hexdigest()
        stream.write(f"{digest}  {path.relative_to(root).as_posix()}\n")
PY
}

fixture_hashes() {
    local root=$1
    local output=$2
    python3 - "$root" "$output" <<'PY'
import hashlib
import pathlib
import sys

root = pathlib.Path(sys.argv[1])
destination = pathlib.Path(sys.argv[2])
with destination.open("w", encoding="utf-8") as stream:
    for path in sorted(root.glob("*.xlsx")):
        stream.write(f"{hashlib.sha256(path.read_bytes()).hexdigest()}  {path.name}\n")
PY
}

fixture_member_hashes() {
    local root=$1
    local output=$2
    python3 - "$root" "$output" <<'PY'
import hashlib
import pathlib
import sys
import zipfile

root = pathlib.Path(sys.argv[1])
destination = pathlib.Path(sys.argv[2])
with destination.open("w", encoding="utf-8") as stream:
    for package in sorted(root.glob("*.xlsx")):
        with zipfile.ZipFile(package) as archive:
            for info in sorted(archive.infolist(), key=lambda item: item.filename):
                digest = hashlib.sha256(archive.read(info)).hexdigest()
                stream.write(f"{digest}  {package.name}!{info.filename}\n")
PY
}

harness_manifest() {
    local output=$1
    python3 - "$script_dir" "$output" <<'PY'
import hashlib
import pathlib
import sys

root = pathlib.Path(sys.argv[1])
destination = pathlib.Path(sys.argv[2])
files = [root / relative for relative in ("Cargo.toml", "Cargo.lock", "src/main.rs", "run.sh")]
with destination.open("w", encoding="utf-8") as stream:
    for path in files:
        if not path.is_file():
            raise SystemExit(f"missing harness file: {path}")
        digest = hashlib.sha256(path.read_bytes()).hexdigest()
        stream.write(f"{digest}  {path.relative_to(root).as_posix()}\n")
PY
}

manifest "$source_root" "$output_dir/source-manifest-before.sha256"
fixture_hashes "$fixture_root" "$output_dir/fixture-hashes-before.sha256"
fixture_member_hashes "$fixture_root" "$output_dir/fixture-member-hashes-before.sha256"
harness_manifest "$output_dir/harness-manifest-before.sha256"
sha256sum "$freeze_receipt" > "$output_dir/freeze-receipt.sha256"

mkdir -p "$work_dir/project/src"
cp "$script_dir/Cargo.toml" "$work_dir/project/Cargo.toml"
cp "$script_dir/Cargo.lock" "$work_dir/project/Cargo.lock"
cp "$script_dir/src/main.rs" "$work_dir/project/src/main.rs"
ln -s "$source_root" "$work_dir/source"
for config in .cargo rust-toolchain.toml clippy.toml; do
    if [[ -e "$source_root/$config" ]]; then
        ln -s "$source_root/$config" "$work_dir/project/$config"
    fi
done
export CARGO_TARGET_DIR="$target_dir"
(cd "$work_dir/project" && cargo build --release --locked --offline --manifest-path Cargo.toml) \
    > "$output_dir/build.log" 2>&1
binary="$target_dir/release/xlsx-form-control-scalar-perf"
[[ -x "$binary" ]] || { printf 'release harness was not built: %s\n' "$binary" >&2; exit 1; }
sha256sum "$binary" > "$output_dir/harness-binary-before.sha256"

cat > "$output_dir/commands.txt" <<EOF
source_root=$source_root
fixture_root=$fixture_root
binary=$binary
warmups=$warmups
iterations=$iterations
EOF

raw="$output_dir/raw-receipts.jsonl"
: > "$raw"
run_fixture() {
    local fixture=$1
    local label=$2
    local expected=$3
    local field=${4:-}
    local field_args=()
    if [[ -n "$field" ]]; then
        field_args=(--field "$field")
    fi
    [[ -f "$fixture_root/$fixture" ]] || { printf 'missing fixture: %s\n' "$fixture_root/$fixture" >&2; exit 2; }
    printf '%s\n' "--fixture $fixture --label $label --expected-controls $expected${field:+ --field $field}" >> "$output_dir/commands.txt"
    "$binary" \
        --fixture "$fixture_root/$fixture" \
        --label "$label" \
        --expected-controls "$expected" \
        "${field_args[@]}" \
        --warmups "$warmups" \
        --iterations "$iterations" \
        >> "$raw"
}
run_fixture singlecontrol.xlsx singlecontrol 1
run_fixture tdf120301_xmlSpaceParsing.xlsx tdf120301_xmlSpaceParsing 2 NoThreeD
run_fixture tdf134769.xlsx tdf134769 1 NoThreeD

sha256sum "$binary" > "$output_dir/harness-binary-after.sha256"
manifest "$source_root" "$output_dir/source-manifest-after.sha256"
fixture_hashes "$fixture_root" "$output_dir/fixture-hashes-after.sha256"
fixture_member_hashes "$fixture_root" "$output_dir/fixture-member-hashes-after.sha256"
harness_manifest "$output_dir/harness-manifest-after.sha256"
cmp "$output_dir/source-manifest-before.sha256" "$output_dir/source-manifest-after.sha256"
cmp "$output_dir/fixture-hashes-before.sha256" "$output_dir/fixture-hashes-after.sha256"
cmp "$output_dir/fixture-member-hashes-before.sha256" "$output_dir/fixture-member-hashes-after.sha256"
cmp "$output_dir/harness-manifest-before.sha256" "$output_dir/harness-manifest-after.sha256"
cmp "$output_dir/harness-binary-before.sha256" "$output_dir/harness-binary-after.sha256"

python3 - "$raw" "$output_dir/run-manifest.json" "$output_dir/receipt-index.json" "$output_dir/sanity-validation.json" "$output_dir/stats.json" "$output_dir/report.md" "$source_root" "$fixture_root" "$freeze_receipt" "$output_dir/source-manifest-before.sha256" "$output_dir/harness-manifest-before.sha256" "$warmups" "$iterations" <<'PY'
import hashlib
import json
import os
import pathlib
import platform
import subprocess
import sys
from collections import Counter, defaultdict

raw_path, manifest_path, index_path, sanity_path, stats_path, report_path, source_root, fixture_root, freeze, source_manifest, harness_manifest_path, warmups, iterations = sys.argv[1:]
warmups = int(warmups)
iterations = int(iterations)
records = []
with pathlib.Path(raw_path).open(encoding="utf-8") as stream:
    for line in stream:
        if line.strip():
            records.append(json.loads(line))
samples = [record for record in records if record.get("record") == "sample"]
correctness = [record for record in records if record.get("record") == "correctness"]
complete = [record for record in records if record.get("record") == "complete"]
lanes = Counter(record["lane"] for record in samples)
fixtures = Counter(record["fixture"] for record in samples)
expected_lanes = {
    "eager_read", "source_read", "eager_noop_save_reopen",
    "source_noop_save_reopen", "eager_scalar_save_reopen",
    "source_scalar_save_reopen", "source_forward_apply", "source_inverse_apply",
}
errors = []
if len(correctness) != 3:
    errors.append(f"expected three correctness records, observed {len(correctness)}")
if len(complete) != 3:
    errors.append(f"expected three completion records, observed {len(complete)}")
if set(lanes) != expected_lanes:
    errors.append(f"lane set mismatch: {sorted(lanes)}")
expected_per_lane = 3 * iterations
expected_per_fixture = 8 * iterations
if any(count != expected_per_lane for count in lanes.values()):
    errors.append(f"each lane must have {expected_per_lane} samples, observed {dict(lanes)}")
if any(count != expected_per_fixture for count in fixtures.values()):
    errors.append(f"each fixture must have {expected_per_fixture} samples, observed {dict(fixtures)}")
for record in correctness:
    for key in ("source_noop_exact", "eager_noop_exact", "source_changed_reopen", "eager_changed_reopen", "inverse_canonical_exact"):
        if record.get(key) is not True:
            errors.append(f"correctness record lacks {key}: {record}")
for record in samples:
    for key in ("elapsed_ns", "alloc_calls", "requested_event_bytes", "peak_live_bytes", "output_checksum"):
        if key not in record:
            errors.append(f"sample lacks {key}: {record}")
    if record.get("lane", "").startswith("source_") and record.get("lane") != "source_forward_apply" and record.get("lane") != "source_inverse_apply":
        if record.get("source_read_calls", 0) < 0:
            errors.append(f"negative source read count: {record}")

def command(command, cwd=None):
    try:
        return subprocess.check_output(command, cwd=cwd, text=True, stderr=subprocess.STDOUT).strip()
    except (OSError, subprocess.CalledProcessError) as error:
        return f"error: {error}"

def memory_total_bytes():
    try:
        for line in pathlib.Path("/proc/meminfo").read_text().splitlines():
            if line.startswith("MemTotal:"):
                return int(line.split()[1]) * 1024
    except (OSError, ValueError, IndexError):
        pass
    return None

freeze_digest = hashlib.sha256(pathlib.Path(freeze).read_bytes()).hexdigest()
source_manifest_digest = hashlib.sha256(pathlib.Path(source_manifest).read_bytes()).hexdigest()
harness_manifest_digest = hashlib.sha256(pathlib.Path(harness_manifest_path).read_bytes()).hexdigest()
compiler = command(["rustc", "-Vv"], cwd=source_root)
target_triple = next(
    (line.split(": ", 1)[1] for line in compiler.splitlines() if line.startswith("host: ")),
    None,
)
manifest = {
    "schema": 1,
    "status": "pass" if not errors else "fail",
    "source_root": source_root,
    "fixture_root": fixture_root,
    "freeze_receipt": freeze,
    "freeze_receipt_sha256": freeze_digest,
    "source_manifest_sha256": source_manifest_digest,
    "harness_manifest_sha256": harness_manifest_digest,
    "compiler": compiler,
    "target_triple": target_triple,
    "cargo": command(["cargo", "-V"], cwd=source_root),
    "toolchain": command(["rustup", "show", "active-toolchain"], cwd=source_root),
    "build": {
        "command": "cargo build --release --locked --offline --manifest-path Cargo.toml",
        "profile": "release",
        "allocator": "std::alloc::System wrapped by TrackingAllocator",
    },
    "host": {
        "uname": platform.platform(),
        "kernel": platform.release(),
        "processor": platform.processor(),
        "machine": platform.machine(),
        "cpu_count": os.cpu_count(),
        "memory_total_bytes": memory_total_bytes(),
        "storage": command(["df", "-P", source_root]),
        "python": platform.python_version(),
    },
    "environment": {
        key: os.environ.get(key)
        for key in ("TMPDIR", "CARGO_TARGET_DIR", "CARGO_HOME", "RUSTFLAGS", "RUSTC_WRAPPER", "RUSTUP_TOOLCHAIN")
    },
    "warmups": warmups,
    "iterations_per_lane": iterations,
    "fixtures": sorted(fixtures),
    "lanes": sorted(lanes),
    "sample_count": len(samples),
    "errors": errors,
}
index = {
    "schema": 1,
    "status": "pass" if not errors else "fail",
    "sample_count": len(samples),
    "correctness_count": len(correctness),
    "complete_count": len(complete),
    "samples_by_lane": dict(sorted(lanes.items())),
    "samples_by_fixture": dict(sorted(fixtures.items())),
    "raw_sha256": hashlib.sha256(pathlib.Path(raw_path).read_bytes()).hexdigest(),
    "errors": errors,
}
sanity = {
    "schema": 1,
    "status": "pass" if not errors else "fail",
    "raw_records": len(records),
    "sample_records": len(samples),
    "correctness_records": len(correctness),
    "errors": errors,
}
pathlib.Path(sanity_path).write_text(json.dumps(sanity, indent=2) + "\n", encoding="utf-8")

def percentile(values, fraction):
    ordered = sorted(values)
    if not ordered:
        return None
    if len(ordered) == 1:
        return ordered[0]
    position = (len(ordered) - 1) * fraction
    lower = int(position)
    upper = min(lower + 1, len(ordered) - 1)
    weight = position - lower
    value = ordered[lower] + (ordered[upper] - ordered[lower]) * weight
    return int(value) if value.is_integer() else value

def metric(values):
    values = [value for value in values if value is not None]
    if not values:
        return None
    return {
        "min": min(values),
        "p50": percentile(values, 0.50),
        "p95": percentile(values, 0.95),
        "p99": percentile(values, 0.99),
        "max": max(values),
    }

metric_names = (
    "elapsed_ns", "alloc_calls", "alloc_bytes_requested", "dealloc_calls",
    "dealloc_bytes_requested", "realloc_calls", "realloc_bytes_requested",
    "requested_event_bytes", "live_bytes_delta", "peak_live_bytes",
    "rss_before_bytes", "rss_after_bytes", "hwm_before_bytes",
    "hwm_after_bytes", "source_read_calls", "source_read_bytes",
    "source_request_bytes", "source_max_request_bytes", "source_len_calls",
    "source_version_calls",
)
grouped = defaultdict(list)
for record in samples:
    grouped[(record["fixture"], record["lane"])].append(record)
groups = []
for (fixture, lane), group in sorted(grouped.items()):
    groups.append({
        "fixture": fixture,
        "lane": lane,
        "sample_count": len(group),
        "field": group[0].get("field"),
        "metrics": {
            name: metric([record.get(name) for record in group])
            for name in metric_names
        },
    })
stats = {
    "schema": 1,
    "status": "pass" if not errors else "fail",
    "sample_count": len(samples),
    "percentile_method": "linear interpolation over sorted raw samples using (n - 1) * fraction",
    "groups": groups,
    "errors": errors,
}
pathlib.Path(stats_path).write_text(json.dumps(stats, indent=2) + "\n", encoding="utf-8")

report_lines = [
    "# XLSX form-control scalar lifecycle performance",
    "",
    f"Status: **{stats['status']}**; raw samples: **{len(samples)}**.",
    "",
    "Percentiles use linear interpolation over each fixture/lane's retained raw samples.",
    "Latency is reported in nanoseconds; allocation and RSS fields retain their raw units.",
    "",
    "| Fixture | Lane | n | Field | elapsed p50 | elapsed p95 | elapsed p99 | peak live bytes p95 | RSS after p95 |",
    "| --- | --- | ---: | --- | ---: | ---: | ---: | ---: | ---: |",
]
for group in groups:
    elapsed = group["metrics"]["elapsed_ns"] or {}
    peak = group["metrics"]["peak_live_bytes"] or {}
    rss = group["metrics"]["rss_after_bytes"] or {}
    report_lines.append(
        f"| {group['fixture']} | {group['lane']} | {group['sample_count']} | "
        f"{group.get('field') or ''} | {elapsed.get('p50', '')} | {elapsed.get('p95', '')} | "
        f"{elapsed.get('p99', '')} | {peak.get('p95', '')} | {rss.get('p95', '')} |"
    )
pathlib.Path(report_path).write_text("\n".join(report_lines) + "\n", encoding="utf-8")
stats_digest = hashlib.sha256(pathlib.Path(stats_path).read_bytes()).hexdigest()
report_digest = hashlib.sha256(pathlib.Path(report_path).read_bytes()).hexdigest()
manifest["stats_sha256"] = stats_digest
manifest["report_sha256"] = report_digest
manifest["stats_path"] = stats_path
manifest["report_path"] = report_path
index["stats_sha256"] = stats_digest
index["report_sha256"] = report_digest
pathlib.Path(manifest_path).write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")
pathlib.Path(index_path).write_text(json.dumps(index, indent=2) + "\n", encoding="utf-8")
if errors:
    raise SystemExit("final-source scalar performance validation failed")
PY

printf 'final-source scalar lifecycle capture complete: %s\n' "$output_dir"
