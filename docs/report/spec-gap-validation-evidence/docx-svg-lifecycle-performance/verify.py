#!/usr/bin/env python3
"""Fail-closed verifier for DOCX SVG lifecycle receipts."""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path

from summarize import LANES, rows


BASELINE = "892441d95db29da4351390716ef5c65b4c7c97de"
EXPECTED_REFUSALS = {"batch_attach_64", "exact_inverse_batch_64"}
U64_MAX = (1 << 64) - 1
PHASE_FIELDS = (
    "capture_ns",
    "stage_ns",
    "commit_ns",
    "publish_ns",
    "reopen_ns",
    "inverse_reopen_ns",
    "inverse_ns",
    "payload_ns",
    "validation_ns",
    "readback_ns",
)
SAMPLE_U64_FIELDS = (
    "elapsed_ns",
    "output_bytes",
    "direct_allocated_bytes",
    "realloc_old_bytes",
    "realloc_new_bytes",
    "deallocated_bytes",
    "requested_alloc_bytes",
    "live_before",
    "live_after",
    "peak_live_delta",
    "allocation_calls",
    "reallocation_calls",
    "deallocation_calls",
    "allocation_failed",
)
SAMPLE_BOOL_FIELDS = (
    "actual_success",
    "semantic_ok",
    "opaque_ok",
    "exact_inverse_ok",
    "lazy_media_cold_ok",
    "alloc_balance_ok",
    "alloc_invalid",
)
MISSING = object()


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def required(mapping: dict[str, object], key: str, context: str) -> object:
    value = mapping.get(key, MISSING)
    require(value is not MISSING, f"{context} missing {key}")
    return value


def unsigned(value: object, label: str) -> int:
    require(type(value) is int, f"{label} must be an integer")
    require(0 <= value <= U64_MAX, f"{label} is outside u64")
    return value


def boolean(value: object, label: str) -> bool:
    require(type(value) is bool, f"{label} must be a boolean")
    return value


def text(value: object, label: str) -> str:
    require(type(value) is str, f"{label} must be a string")
    return value


def decimal_u64(value: str, label: str, *, positive: bool = False) -> int:
    require(type(value) is str, f"{label} must be text")
    require(value != "" and value.isascii() and value.isdecimal(), f"{label} is not an unsigned integer")
    parsed = int(value, 10)
    require(parsed <= U64_MAX, f"{label} is outside u64")
    require(not positive or parsed > 0, f"{label} must be positive")
    return parsed


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def verify_manifest(path: Path, evidence: Path) -> int:
    lines = path.read_text().splitlines()
    require(lines and lines[0] == "format=docx-svg-lifecycle-build-source-v1", "manifest format changed")
    fields = dict(line.split("=", 1) for line in lines[1:] if "=" in line and not line.startswith("file="))
    require(fields.get("git_head") == BASELINE, "manifest baseline changed")
    require(fields.get("baseline_commit") == BASELINE, "manifest commit field changed")
    require(fields.get("source_snapshot") == "baseline-source", "source snapshot path changed")
    checked = 0
    for line in lines[1:]:
        if not line.startswith("file="):
            continue
        shown, expected = line[len("file=") :].split("\t", 1)
        current = evidence / shown
        require(current.is_file(), f"manifest input missing: {shown}")
        require(sha256(current) == expected, f"manifest input changed: {shown}")
        checked += 1
    require(checked >= 14, f"too few source and harness files frozen: {checked}")
    return checked


def verify_corpus(path: Path, native_root: Path) -> None:
    corpus = json.loads(path.read_text())
    require(corpus["baseline_commit"] == BASELINE, "corpus baseline changed")
    for fixture in corpus["fixtures"]:
        source = native_root / fixture["source_path"]
        copy = path.parent / fixture["copy_path"]
        require(source.is_file(), f"native source fixture missing: {source}")
        require(copy.is_file(), f"copied fixture missing: {copy}")
        require(source.stat().st_size == fixture["bytes"], f"native fixture size changed: {source}")
        require(copy.stat().st_size == fixture["bytes"], f"copied fixture size changed: {copy}")
        require(sha256(source) == fixture["sha256"], f"native fixture hash changed: {source}")
        require(sha256(copy) == fixture["sha256"], f"copied fixture hash changed: {copy}")


def verify_timing(path: Path, stderr: Path) -> None:
    require(stderr.is_file(), f"stderr receipt missing: {stderr}")
    require(stderr.read_bytes() == b"", f"unexpected profile stderr: {stderr}")
    require(path.is_file(), f"RSS receipt missing: {path}")
    statuses = [
        line.strip()
        for line in path.read_text().splitlines()
        if line.strip().startswith("Exit status:")
    ]
    require(statuses == ["Exit status: 0"], f"profile process failed: {path}")
    marker = "Maximum resident set size (kbytes):"
    rss_lines = [
        line.split(":", 1)[1].strip()
        for line in path.read_text().splitlines()
        if line.lstrip().startswith(marker)
    ]
    require(len(rss_lines) == 1, f"RSS receipt is missing or duplicated: {path}")
    decimal_u64(rss_lines[0], f"RSS in {path}", positive=True)


def verify_sample(sample: dict[str, object], lane: str, expected_success: bool, path: Path) -> None:
    require(type(sample) is dict, f"sample must be an object: {path}")
    require(type(expected_success) is bool, f"expected success must be boolean: {path}")
    unsigned_values = {
        key: unsigned(required(sample, key, "sample"), f"{path}:{key}")
        for key in SAMPLE_U64_FIELDS
    }
    for key in SAMPLE_BOOL_FIELDS:
        boolean(required(sample, key, "sample"), f"{path}:{key}")
    require(
        unsigned_values["elapsed_ns"] > 0,
        f"elapsed_ns must be positive: {path}",
    )
    require(
        required(sample, "actual_success", "sample") is expected_success,
        f"success status mismatch: {path}",
    )
    for gate in ("semantic_ok", "opaque_ok", "lazy_media_cold_ok", "exact_inverse_ok"):
        require(required(sample, gate, "sample") is True, f"{gate} failed: {path}")

    phases = required(sample, "phases", "sample")
    require(type(phases) is dict, f"phases must be an object: {path}")
    require(set(phases) == set(PHASE_FIELDS), f"phase keys changed: {path}")
    phase_values = {
        key: unsigned(required(phases, key, "phases"), f"{path}:phases.{key}")
        for key in PHASE_FIELDS
    }
    elapsed_ns = unsigned_values["elapsed_ns"]
    for key, value in phase_values.items():
        require(value <= elapsed_ns, f"phase exceeds elapsed_ns ({key}): {path}")
    require(sum(phase_values.values()) <= elapsed_ns, f"phase sum exceeds elapsed_ns: {path}")

    direct = unsigned_values["direct_allocated_bytes"]
    new = unsigned_values["realloc_new_bytes"]
    old = unsigned_values["realloc_old_bytes"]
    freed = unsigned_values["deallocated_bytes"]
    live_before = unsigned_values["live_before"]
    live_after = unsigned_values["live_after"]
    expected_live = live_before + direct + new - old - freed
    require(0 <= expected_live <= U64_MAX, f"live-byte equation overflows: {path}")
    require(unsigned_values["requested_alloc_bytes"] == direct + new, f"allocation equation failed: {path}")
    require(unsigned_values["live_after"] == expected_live, f"live-byte equation failed: {path}")
    require(
        unsigned_values["peak_live_delta"] >= max(0, live_after - live_before),
        f"peak live delta is below live-byte increase: {path}",
    )
    require(required(sample, "alloc_balance_ok", "sample") is True, f"allocator balance failed: {path}")
    require(required(sample, "alloc_invalid", "sample") is False, f"allocator underflow failed: {path}")
    require(unsigned_values["allocation_failed"] == 0, f"allocator failure failed: {path}")

    error = required(sample, "error", "sample")
    physical_readback = required(sample, "source_readback_physical_ok", "sample")
    metadata_readback = required(sample, "source_readback_metadata_ok", "sample")
    if expected_success:
        require(error is None, f"unexpected refusal: {path}")
        require(unsigned_values["output_bytes"] > 0, f"successful lane emitted no output: {path}")
        require(physical_readback is None, f"unexpected refusal readback gate: {path}")
        require(metadata_readback is None, f"unexpected refusal metadata gate: {path}")
    else:
        require(lane in EXPECTED_REFUSALS, f"unclassified refusal lane: {lane}")
        require(type(error) is dict, f"refusal receipt missing: {path}")
        require(set(error) == {"class", "message", "typed_match"}, f"refusal receipt schema changed: {path}")
        require(text(required(error, "class", "error"), f"{path}:error.class") == "topology_part_bound", f"refusal class changed: {path}")
        text(required(error, "message", "error"), f"{path}:error.message")
        require(boolean(required(error, "typed_match", "error"), f"{path}:error.typed_match") is True, f"typed refusal match missing: {path}")
        require(physical_readback is True, f"physical refusal readback failed: {path}")
        require(metadata_readback is True, f"metadata refusal readback failed: {path}")
        require(unsigned_values["output_bytes"] == 0, f"refusal emitted output: {path}")


def verify_receipts(results: Path, mode: str) -> tuple[int, int]:
    expected_processes = 1 if mode == "smoke" else 3
    expected_samples = 1 if mode == "smoke" else 20
    total_samples = 0
    total_processes = 0
    for lane in LANES:
        expected_success = lane not in EXPECTED_REFUSALS
        paths = [
            results / f"{mode}-{lane}-p{process}.json"
            for process in range(1, expected_processes + 1)
            if (results / f"{mode}-{lane}-p{process}.json").is_file()
        ]
        require(len(paths) == expected_processes, f"process count mismatch: {lane}")
        hashes: set[str] = set()
        fnv_hashes: set[int] = set()
        sizes: set[int] = set()
        for path in paths:
            value = json.loads(path.read_text())
            require(type(value) is dict, f"receipt must be an object: {path}")
            require(text(required(value, "schema", "receipt"), f"{path}:schema") == "docx-svg-lifecycle-profile-v1", f"schema mismatch: {path}")
            require(text(required(value, "baseline_commit", "receipt"), f"{path}:baseline_commit") == BASELINE, f"baseline mismatch: {path}")
            require(text(required(value, "opc_baseline_label", "receipt"), f"{path}:opc_baseline_label") == "OPC57680dc86", f"OPC baseline mismatch: {path}")
            require(text(required(value, "lane", "receipt"), f"{path}:lane") == lane, f"lane mismatch: {path}")
            receipt_expected_success = boolean(required(value, "expected_success", "receipt"), f"{path}:expected_success")
            require(receipt_expected_success is expected_success, f"expected status mismatch: {path}")
            unsigned(required(value, "warmup", "receipt"), f"{path}:warmup")
            sample_count = unsigned(required(value, "sample_count", "receipt"), f"{path}:sample_count")
            samples = required(value, "samples", "receipt")
            require(type(samples) is list, f"samples must be an array: {path}")
            require(sample_count == len(samples) == expected_samples, f"sample count mismatch: {path}")
            input_sha = text(required(value, "input_sha256", "receipt"), f"{path}:input_sha256")
            require(
                len(input_sha) == 64 and input_sha.isascii() and all(char in "0123456789abcdef" for char in input_sha),
                f"input SHA-256 is malformed: {path}",
            )
            input_fnv = unsigned(required(value, "input_hash_fnv1a64", "receipt"), f"{path}:input_hash_fnv1a64")
            input_bytes = unsigned(required(value, "input_bytes", "receipt"), f"{path}:input_bytes")
            unsigned(required(value, "owners", "receipt"), f"{path}:owners")
            text(required(value, "fixture_kind", "receipt"), f"{path}:fixture_kind")
            require(boolean(required(value, "source_backed_api", "receipt"), f"{path}:source_backed_api") is True, f"source-backed API gate failed: {path}")
            hashes.add(input_sha)
            fnv_hashes.add(input_fnv)
            sizes.add(input_bytes)
            timing = path.with_suffix(".time.txt")
            verify_timing(timing, path.with_suffix(".stderr.log"))
            for sample in samples:
                verify_sample(sample, lane, expected_success, path)
        require(len(hashes) == len(fnv_hashes) == len(sizes) == 1, f"fixture identity changed: {lane}")
        total_samples += expected_samples * expected_processes
        total_processes += expected_processes
    return total_processes, total_samples


def verify_report(path: Path, table: list[dict[str, object]]) -> None:
    parsed: dict[str, dict[str, object]] = {}
    for line in path.read_text().splitlines():
        if not line.startswith("| ") or line.startswith("|---") or line.startswith("| lane"):
            continue
        fields = [field.strip() for field in line.strip().strip("|").split("|")]
        require(len(fields) == 9, f"malformed report row: {line}")
        elapsed = tuple(int(item) for item in fields[4].split("/"))
        alloc = tuple(int(item) for item in fields[5].split("/"))
        peak = tuple(int(item) for item in fields[6].split("/"))
        rss_values = tuple(int(item) for item in fields[7].replace("–", "-").split("-"))
        dominant, p50 = fields[8].rsplit(" (", 1)
        parsed[fields[0]] = {
            "processes": int(fields[1]),
            "samples": int(fields[2]),
            "input_bytes": int(fields[3]),
            "elapsed": elapsed,
            "alloc": alloc,
            "peak": peak,
            "rss": rss_values,
            "dominant_phase": dominant,
            "dominant_p50": int(p50.rstrip(")")),
        }
    require(set(parsed) == {str(row["lane"]) for row in table}, "report lane set differs")
    for row in table:
        actual = parsed[str(row["lane"])]
        for field in ("processes", "samples", "input_bytes", "elapsed", "alloc", "peak", "rss", "dominant_phase"):
            require(actual[field] == row[field], f"report {field} differs: {row['lane']}")
        require(actual["dominant_p50"] == row["phase_p50"][row["dominant_phase"]], f"report phase differs: {row['lane']}")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--evidence", type=Path, required=True)
    parser.add_argument("--native-root", type=Path, required=True)
    parser.add_argument("--results", type=Path, required=True)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--mode", choices=("smoke", "full"), required=True)
    parser.add_argument("--report", type=Path, required=True)
    parser.add_argument("--corpus", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    before = args.manifest
    after = args.results / f"{args.mode}-source-manifest-after.txt"
    require(before.read_bytes() == after.read_bytes(), "source manifest changed during profile")
    require(
        (args.results / f"{args.mode}-metadata-before.json").read_bytes()
        == (args.results / f"{args.mode}-metadata-after.json").read_bytes(),
        "Cargo metadata changed",
    )
    require(
        (args.results / f"{args.mode}-binary.sha256").read_bytes()
        == (args.results / f"{args.mode}-binary-after.sha256").read_bytes(),
        "profile binary changed",
    )
    manifest_inputs = verify_manifest(before, args.evidence)
    verify_corpus(args.corpus, args.native_root)
    processes, samples = verify_receipts(args.results, args.mode)
    table = rows(args.results, args.mode)
    verify_report(args.report, table)
    result = {
        "passed": True,
        "mode": args.mode,
        "lanes": len(LANES),
        "processes": processes,
        "samples": samples,
        "manifest_inputs_checked": manifest_inputs,
        "expected_refusals": sorted(EXPECTED_REFUSALS),
        "allocation_peak_rss_separate": True,
        "report_recomputed_from_raw_receipts": True,
        "fail_closed_gates": [
            "source_snapshot",
            "source_manifest",
            "fixture_hash",
            "semantic",
            "opaque",
            "lazy_media",
            "exact_inverse",
            "typed_error",
            "failure_physical_readback",
            "failure_metadata_readback",
            "allocator",
            "rss",
            "process_exit",
        ],
    }
    args.output.write_text(json.dumps(result, indent=2) + "\n")
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
