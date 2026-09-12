#!/usr/bin/env bash
set -euo pipefail

script_dir=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
evidence_dir=$script_dir
repo_root=${1:-${LITCHI_FORM_OWNER_BASE_REPO:-}}
if [[ -z "$repo_root" ]]; then
    printf 'usage: %s BASE_REPOSITORY [OUTPUT_DIRECTORY]\n' "$0" >&2
    exit 2
fi
repo_root=$(CDPATH= cd -- "$repo_root" && pwd)

if [[ -n "${2:-}" ]]; then
    output_dir=$2
    mkdir -p "$output_dir"
    output_dir=$(CDPATH= cd -- "$output_dir" && pwd)
    if find "$output_dir" -mindepth 1 -print -quit | grep -q .; then
        printf 'refusing to overwrite non-empty output directory: %s\n' "$output_dir" >&2
        exit 2
    fi
else
    output_dir=$(mktemp -d "${TMPDIR:-/tmp}/form-owner-scan-replay.XXXXXX")
fi

run_root=$(mktemp -d "${TMPDIR:-/tmp}/form-owner-scan-replay-root.XXXXXX")
cleanup() { rm -rf "$run_root"; }
trap cleanup EXIT

before_source="$run_root/baseline/source"
after_source="$run_root/candidate/source"
"$evidence_dir/restore-source.sh" "$repo_root" "$before_source" baseline \
    >"$output_dir/restore-baseline.log"
"$evidence_dir/restore-source.sh" "$repo_root" "$after_source" candidate \
    >"$output_dir/restore-candidate.log"

fixture_dir="$run_root/fixtures"
mkdir -p "$fixture_dir"
cp "$evidence_dir/fixtures/"*.xlsx "$fixture_dir/"

for mode in baseline candidate; do
    source_root="$run_root/$mode/source"
    package_root="$run_root/$mode/package"
    target_dir="$run_root/$mode/target"
    mkdir -p "$package_root/src"
    cp "$evidence_dir/harness/Cargo.toml" "$package_root/Cargo.toml"
    cp "$evidence_dir/harness/Cargo.lock" "$package_root/Cargo.lock"
    cp "$evidence_dir/harness/src/main.rs" "$package_root/src/main.rs"
    if [[ -d "$source_root/.cargo" ]]; then
        ln -s "$source_root/.cargo" "$package_root/.cargo"
    fi
    CARGO_TARGET_DIR="$target_dir" cargo build --release --locked --offline \
        --manifest-path "$package_root/Cargo.toml" \
        >"$output_dir/build-${mode}.log" 2>&1
    binary="$target_dir/release/form-owner-perf-harness"
    sha256sum "$binary" >"$output_dir/harness-binary-${mode}.sha256"
    raw="$output_dir/${mode}-receipts-replay.jsonl"
    : >"$raw"
    run() {
        local label=$1 path=$2 count=$3
        "$binary" --fixture "$path" --label "$label" \
            --expected-count "$count" --iterations 7 >>"$raw"
    }
    run button-small "$fixture_dir/button-form-control.xlsx" 1
    run tdf60673-medium "$fixture_dir/tdf60673.xlsx" 2
    run many-controls-16 "$fixture_dir/many-controls-16.xlsx" 16
    run diagnostic-mirror "$fixture_dir/diagnostic-mirror.xlsx" 2
    "$binary" --fixture "$fixture_dir/foreign-wrapper-refusal.xlsx" \
        --label foreign-wrapper-refusal --iterations 7 --expect-refusal >>"$raw"
done

python3 - "$output_dir" "$evidence_dir" "$run_root" <<'PY'
import hashlib
import json
import pathlib
import subprocess
import sys

output, evidence, run_root = map(pathlib.Path, sys.argv[1:])

def sha256(path):
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()

def jsonl_summary(path):
    rows = [json.loads(line) for line in path.read_text().splitlines()]
    correctness = [row for row in rows if row["record"] in {"correctness", "refusal_correctness"}]
    receipts = [row for row in rows if row["record"] == "receipt"]
    if len(correctness) != 5 or len(receipts) != 294:
        raise SystemExit(f"unexpected {path.name} counts: correctness={len(correctness)} receipts={len(receipts)}")
    return {
        "sha256": sha256(path),
        "correctness_records": len(correctness),
        "receipt_records": len(receipts),
        "fixtures": sorted({row["fixture"] for row in receipts}),
        "lanes": sorted({row["lane"] for row in receipts}),
        "iterations_per_fixture": 7,
    }

def jsonl_rows(path, record):
    return [
        json.loads(line)
        for line in path.read_text().splitlines()
        if json.loads(line)["record"] == record
    ]

baseline_rows = jsonl_rows(output / "baseline-receipts-replay.jsonl", "receipt")
candidate_rows = jsonl_rows(output / "candidate-receipts-replay.jsonl", "receipt")
sort_key = lambda row: (row["fixture"], row["lane"], row["iteration"])
baseline_rows.sort(key=sort_key)
candidate_rows.sort(key=sort_key)
if len(baseline_rows) != len(candidate_rows):
    raise SystemExit("baseline/candidate receipt counts differ")
source_fields = (
    "fixture",
    "lane",
    "iteration",
    "source_read_calls",
    "source_read_bytes",
    "source_request_bytes",
    "source_max_request_bytes",
    "source_len_calls",
    "source_version_calls",
)
source_metrics_identical = all(
    tuple(before[field] for field in source_fields)
    == tuple(after[field] for field in source_fields)
    for before, after in zip(baseline_rows, candidate_rows)
)
if not source_metrics_identical:
    raise SystemExit("baseline/candidate source read metrics differ")

baseline_correctness = jsonl_rows(output / "baseline-receipts-replay.jsonl", "correctness")
candidate_correctness = jsonl_rows(output / "candidate-receipts-replay.jsonl", "correctness")
baseline_refusal = jsonl_rows(output / "baseline-receipts-replay.jsonl", "refusal_correctness")
candidate_refusal = jsonl_rows(output / "candidate-receipts-replay.jsonl", "refusal_correctness")
if baseline_correctness != candidate_correctness or baseline_refusal != candidate_refusal:
    raise SystemExit("baseline/candidate correctness records differ")

def owner_hash(path):
    return subprocess.check_output(["git", "hash-object", str(path)], text=True).strip()

metadata = {
    "status": "pass",
    "base_commit": (evidence / "base-commit.txt").read_text().strip(),
    "source_patch_sha256": sha256(evidence / "source-delta.patch"),
    "source_lock_sha256": sha256(evidence / "source-Cargo.lock"),
    "harness_lock_sha256": sha256(evidence / "harness/Cargo.lock"),
    "fixture_hashes": {path.name: sha256(path) for path in sorted((evidence / "fixtures").glob("*.xlsx"))},
    "replayed": {
        "baseline": jsonl_summary(output / "baseline-receipts-replay.jsonl"),
        "candidate": jsonl_summary(output / "candidate-receipts-replay.jsonl"),
    },
    "receipt_identity": {
        "correctness_records_identical": True,
        "source_read_metrics_identical": True,
        "timing_and_allocator_fields_excluded": True,
    },
    "source_bindings": {
        "baseline_owner_git_blob": owner_hash(run_root / "baseline/source/crates/litchi-xlsx/src/form_control/owner.rs"),
        "candidate_owner_git_blob": owner_hash(run_root / "candidate/source/crates/litchi-xlsx/src/form_control/owner.rs"),
        "baseline_owner_sha256": sha256(run_root / "baseline/source/crates/litchi-xlsx/src/form_control/owner.rs"),
        "candidate_owner_sha256": sha256(run_root / "candidate/source/crates/litchi-xlsx/src/form_control/owner.rs"),
    },
    "historical_final_receipts": {
        "baseline": sha256(evidence / "receipts/baseline-final.jsonl"),
        "candidate": sha256(evidence / "receipts/candidate-final.jsonl"),
    },
    "speedup_claim": False,
}
expected = {
    "baseline_owner_git_blob": "5ee687c00133761fddc6ee3cd4be3cc9ebea8e67",
    "candidate_owner_git_blob": "541c1a3ca4d4b742ac6a134aae8dc942f61404ed",
    "baseline_owner_sha256": "ac8f6dfadd81d1a584488d0dcb3fcb3ef6ceaba75d51f99dc9fab0843df0bb3a",
    "candidate_owner_sha256": "778c510def4174e196ae277ec92127337f49bb28b8834f9bbcdb4b889b60d5c4",
}
for key, value in expected.items():
    if metadata["source_bindings"][key] != value:
        raise SystemExit(f"{key} mismatch: {metadata['source_bindings'][key]} != {value}")
(output / "replay-metadata.json").write_text(json.dumps(metadata, indent=2) + "\n")
PY

printf 'replay output: %s\n' "$output_dir"
