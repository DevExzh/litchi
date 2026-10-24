#!/usr/bin/env python3
"""Validate the reusable RTF and OLE/XLS native probe receipts.

The Rust examples own their route-level output, error, replay and fidelity
assertions.  This validator owns the paired-run boundary: exact baseline
identity, binary/lock/probe hashes, fixed ABBA order, quiet-process/cap
receipts, and complete route output streams.  It never interprets elapsed
values as an acceptance result.  Allocation instrumentation is intentionally
kept out of the normal timing binaries; its separate, non-claimable phase is
recorded by the protocol document.
"""

from __future__ import annotations

import csv
import hashlib
import json
import math
from pathlib import Path
import re
import sys
from typing import Any, Iterable


SCHEMA = "litchi.native-qualification.v1"
SUMMARY_SCHEMA = "litchi.legacy-native-qualification.v1"
ABBA_ORDER = ("before", "after", "after", "before")
TARGET_PRE_FLIGHT = 19_200_000_000
TARGET_HARD = 20_000_000_000
MEMORY_HARD = 20_000_000_000
FULL_SHA = re.compile(r"^[0-9a-f]{40}$")
HASH = re.compile(r"^[0-9a-f]{64}$")

CONTRACTS: dict[str, dict[str, Any]] = {
    "legacy-rtf": {
        "baseline": "86d708878fa537eab42fd4bd5ffbc0f79c606191",
        "blocks": 5,
        "programs": ("routes",),
        "routes": (
            "open_1k",
            "noop_1k",
            "text_commit_1k",
            "text_commit_save_1k",
            "replay_1k",
            "tail_commit_1k",
            "unsupported_content",
            "protected_commit",
        ),
        "probe_contract": "internal route output/error/refusal/fidelity assertions",
    },
    "legacy-ole-properties": {
        "baseline": "604822b44107999aaddb0bffb96b705e9464f546",
        "blocks": 1,
        "programs": ("read", "edit"),
        "routes": ("read", "edit"),
        "samples": 500,
        "probe_contract": "six-property read/edit fixture with exact section output checks",
    },
    "legacy-xls-properties": {
        "baseline": "509bf28e7531ad75f278eaec3a732714a243915a",
        "blocks": 1,
        "programs": ("metadata",),
        "routes": ("doc", "xls", "ppt"),
        "probe_contract": "ole_metadata_bench native/unified byte-equivalence and XLS metadata edit checks",
    },
}

# The XLS source-parent ref is derived from the lazy-entry record.  Keeping it
# in one place makes an accidental substitution of a later lazy source fail
# closed.  The record remains the authority for the derivation rationale.
class ValidationError(ValueError):
    """A receipt or route contract failed."""


def _expect(condition: bool, message: str) -> None:
    if not condition:
        raise ValidationError(message)


def _sha256(path: Path) -> tuple[str, int]:
    _expect(not path.is_symlink(), f"symlink is forbidden: {path}")
    _expect(path.is_file(), f"file is missing: {path}")
    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        while block := stream.read(1024 * 1024):
            digest.update(block)
            size += len(block)
    return digest.hexdigest(), size


def _read_json(path: Path) -> dict[str, Any]:
    _expect(not path.is_symlink() and path.is_file(), f"JSON file is missing or symlinked: {path}")
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise ValidationError(f"cannot read JSON {path}: {error}") from error
    _expect(isinstance(value, dict), f"JSON root is not an object: {path}")
    return value


def _read_events(path: Path) -> list[dict[str, Any]]:
    _expect(not path.is_symlink() and path.is_file(), f"event journal is missing or symlinked: {path}")
    records: list[dict[str, Any]] = []
    for line_number, line in enumerate(path.read_text(encoding="utf-8").splitlines(), start=1):
        try:
            value = json.loads(line)
        except json.JSONDecodeError as error:
            raise ValidationError(f"invalid event JSON at {path}:{line_number}: {error}") from error
        _expect(isinstance(value, dict), f"event at {path}:{line_number} is not an object")
        records.append(value)
    return records


def _validate_hash_record(record: Any, label: str, *, require_nonempty: bool = True) -> dict[str, Any]:
    _expect(isinstance(record, dict), f"{label} is not an object")
    path = Path(str(record.get("path", "")))
    digest, size = _sha256(path)
    _expect(record.get("sha256") == digest, f"{label} SHA-256 changed: {path}")
    _expect(record.get("bytes") == size, f"{label} byte count changed: {path}")
    _expect(HASH.fullmatch(digest) is not None and (size > 0 or not require_nonempty), f"{label} has an invalid identity")
    return {"path": str(path), "bytes": size, "sha256": digest}


def _validate_build_receipt(
    output_dir: Path,
    lane: str,
    contract: dict[str, Any],
    baseline: str,
    candidate: str,
    before_target: Path,
    after_target: Path,
) -> tuple[dict[str, Any], dict[tuple[int, int, str], dict[str, Any]]]:
    receipt_path = output_dir / "build-receipt.json"
    receipt = _read_json(receipt_path)
    _expect(receipt.get("schema") == SCHEMA, "build receipt schema mismatch")
    _expect(receipt.get("status") == "passed" and receipt.get("phase") == "build", "build receipt is not passed")
    _expect(receipt.get("lane") == lane, "build receipt lane mismatch")
    _expect(receipt.get("abba_order") == list(ABBA_ORDER), "build receipt ABBA order mismatch")
    _expect(receipt.get("abba_blocks", 1) == contract["blocks"], "build receipt block count mismatch")
    _expect(
        receipt.get("limits")
        == {
            "target_preflight_bytes": TARGET_PRE_FLIGHT,
            "target_hard_bytes": TARGET_HARD,
            "memory_bytes": MEMORY_HARD,
        },
        "build receipt limits are not the fixed qualification caps",
    )
    identities = receipt.get("identities")
    _expect(isinstance(identities, dict), "build receipt identities are missing")
    expected_heads = {"before": baseline, "after": candidate}
    expected_targets = {"before": before_target.resolve(), "after": after_target.resolve()}
    identity_summary: dict[str, Any] = {}
    for leg, expected_head in expected_heads.items():
        identity = identities.get(leg)
        _expect(isinstance(identity, dict), f"build receipt identity is missing for {leg}")
        _expect(identity.get("head") == expected_head, f"{leg} source HEAD is not the reviewed baseline/candidate")
        _expect(Path(str(identity.get("target_dir", ""))).resolve() == expected_targets[leg], f"{leg} target identity mismatch")
        lock_paths = identity.get("lock_paths")
        probe_paths = identity.get("probe_paths")
        _expect(isinstance(lock_paths, list) and lock_paths, f"{leg} lock identity is missing")
        _expect(isinstance(probe_paths, list) and probe_paths, f"{leg} probe identity is missing")
        files = identity.get("identity_files")
        _expect(isinstance(files, list) and files, f"{leg} source identity files are missing")
        normalized_files = [_validate_hash_record(item, f"{leg} identity file") for item in files]
        file_paths = {item["path"] for item in normalized_files}
        _expect(set(lock_paths).issubset(file_paths), f"{leg} lock path is not bound into identity")
        _expect(set(probe_paths).issubset(file_paths), f"{leg} probe path is not bound into identity")
        artifacts = identity.get("artifacts")
        _expect(isinstance(artifacts, list) and artifacts, f"{leg} binary artifacts are missing")
        normalized_artifacts = [_validate_hash_record(item, f"{leg} binary artifact") for item in artifacts]
        for item in normalized_artifacts:
            path = Path(item["path"])
            _expect(path.parent != path and path.is_file(), f"{leg} artifact is not a file: {path}")
        identity_summary[leg] = {
            "head": expected_head,
            "tree": identity.get("tree"),
            "source_identity_sha256": identity.get("identity_sha256"),
            "lock_paths": [str(path) for path in lock_paths],
            "probe_paths": [str(path) for path in probe_paths],
            "identity_files": normalized_files,
            "artifacts": normalized_artifacts,
        }

    event_path = output_dir / "events-abba.jsonl"
    events = _read_events(event_path)
    expected_events: list[tuple[int, int, str, str]] = []
    for block in range(1, int(contract["blocks"]) + 1):
        for index, leg in enumerate(ABBA_ORDER, start=1):
            for program in contract["programs"]:
                expected_events.append((block, index, leg, program))
    _expect(len(events) == len(expected_events), "ABBA event count differs from the fixed workload")
    event_map: dict[tuple[int, int, str], dict[str, Any]] = {}
    expected_heads = {"before": baseline, "after": candidate}
    for event, (block, index, leg, program) in zip(events, expected_events):
        expected_id = (
            f"abba-{index:02d}-{leg}-{program}"
            if contract["blocks"] == 1
            else f"abba-block-{block:02d}-{index:02d}-{leg}-{program}"
        )
        _expect(event.get("event_id") == expected_id, f"unexpected ABBA event order for {program}")
        _expect(event.get("schema") == SCHEMA and event.get("leg") == leg, f"ABBA event identity mismatch for {program}")
        _expect(event.get("source_head") == expected_heads[leg], f"ABBA source identity mismatch for {program}")
        _expect(event.get("returncode") == 0 and event.get("runner_error") is None, f"ABBA child failed for {program}")
        _expect(not event.get("process_guard_violations"), f"ABBA host was not quiet for {program}")
        _expect(not event.get("post_exit_descendants") and not event.get("post_exit_offenders"), f"ABBA descendants/offenders remained for {program}")
        environment = event.get("environment")
        _expect(isinstance(environment, dict), f"ABBA environment missing for {program}")
        _expect(
            environment.get("CARGO_BUILD_JOBS") == "1"
            and environment.get("CARGO_INCREMENTAL") == "0"
            and environment.get("LITCHI_MEM_MAX") == "20G",
            f"ABBA caps differ for {program}",
        )
        target = event.get("target")
        _expect(isinstance(target, dict), f"ABBA target receipt missing for {program}")
        _expect(
            Path(str(target.get("path", ""))).resolve()
            == {"before": before_target.resolve(), "after": after_target.resolve()}[leg],
            f"ABBA target identity mismatch for {program}",
        )
        _expect(
            isinstance(target.get("bytes_after"), int)
            and 0 <= target["bytes_after"] <= TARGET_HARD,
            f"ABBA target cap failed for {program}",
        )
        stdout = _validate_hash_record(event.get("stdout"), f"{program} stdout")
        _validate_hash_record(event.get("stderr"), f"{program} stderr", require_nonempty=False)
        event_map[(block, index, program)] = {"event": event, "stdout": Path(stdout["path"])}
    return {"receipt": receipt, "identities": identity_summary}, event_map


def _validate_rtf_log(path: Path, routes: Iterable[str]) -> dict[str, Any]:
    _expect(not path.is_symlink() and path.is_file(), f"RTF report is missing or symlinked: {path}")
    rows: list[dict[str, Any]] = []
    for line_number, line in enumerate(path.read_text(encoding="utf-8").splitlines(), start=1):
        try:
            value = json.loads(line)
        except json.JSONDecodeError as error:
            raise ValidationError(f"invalid RTF JSON at {path}:{line_number}: {error}") from error
        _expect(isinstance(value, dict), f"RTF row is not an object: {path}:{line_number}")
        _expect(value.get("scenario") in routes, f"unexpected RTF route in {path}: {value.get('scenario')!r}")
        median = value.get("median_us")
        _expect(isinstance(median, (int, float)) and math.isfinite(float(median)) and median > 0, f"invalid RTF median in {path}")
        rows.append({"scenario": value["scenario"], "median_us": float(median)})
    expected = list(routes)
    _expect([row["scenario"] for row in rows] == expected, f"RTF route/output set differs in {path}")
    return {"path": str(path), "bytes": _sha256(path)[1], "sha256": _sha256(path)[0], "routes": expected}


def _validate_csv_log(path: Path, route: str, samples: int) -> dict[str, Any]:
    _expect(not path.is_symlink() and path.is_file(), f"CSV report is missing or symlinked: {path}")
    with path.open(newline="", encoding="utf-8") as stream:
        rows = list(csv.reader(stream))
    _expect(rows and rows[0] == ["route", "sample", "nanoseconds", "output_bytes"], f"CSV header mismatch in {path}")
    data = rows[1:]
    _expect(len(data) == samples, f"CSV sample count differs in {path}")
    for expected_sample, row in enumerate(data):
        _expect(len(row) == 4 and row[0] == route, f"CSV route shape differs in {path}")
        try:
            sample, elapsed, output_bytes = int(row[1]), int(row[2]), int(row[3])
        except ValueError as error:
            raise ValidationError(f"CSV numeric field is invalid in {path}") from error
        _expect(sample == expected_sample and elapsed >= 0 and output_bytes > 0, f"CSV sample/output check failed in {path}")
    digest, size = _sha256(path)
    return {"path": str(path), "bytes": size, "sha256": digest, "route": route, "samples": samples}


def _validate_metadata_log(path: Path, cases: Iterable[str]) -> dict[str, Any]:
    """Validate the existing metadata benchmark's three deterministic rows."""

    _expect(not path.is_symlink() and path.is_file(), f"metadata report is missing or symlinked: {path}")
    rows: list[dict[str, Any]] = []
    for line_number, line in enumerate(path.read_text(encoding="utf-8").splitlines(), start=1):
        try:
            value = json.loads(line)
        except json.JSONDecodeError as error:
            raise ValidationError(f"invalid metadata JSON at {path}:{line_number}: {error}") from error
        _expect(isinstance(value, dict), f"metadata row is not an object: {path}:{line_number}")
        _expect(value.get("case") in cases, f"unexpected metadata case in {path}: {value.get('case')!r}")
        _expect(isinstance(value.get("bytes"), int) and value["bytes"] > 0, f"invalid metadata bytes in {path}")
        _expect(isinstance(value.get("repeats"), int) and value["repeats"] > 0, f"invalid metadata repeats in {path}")
        for field in ("native_us", "unified_us", "overhead_pct"):
            number = value.get(field)
            _expect(isinstance(number, (int, float)) and math.isfinite(float(number)), f"invalid metadata {field} in {path}")
        _expect(value["native_us"] > 0 and value["unified_us"] > 0, f"metadata elapsed values are not positive in {path}")
        rows.append({"case": value["case"], "bytes": value["bytes"], "repeats": value["repeats"]})
    expected = list(cases)
    _expect([row["case"] for row in rows] == expected, f"metadata route/output set differs in {path}")
    digest, size = _sha256(path)
    return {"path": str(path), "bytes": size, "sha256": digest, "cases": expected}


def validate_directory(
    output_dir: Path,
    lane: str,
    baseline: str,
    candidate: str,
    before_target: Path,
    after_target: Path,
) -> dict[str, Any]:
    _expect(lane in CONTRACTS, f"unknown legacy native lane: {lane}")
    contract = CONTRACTS[lane]
    _expect(baseline == contract["baseline"], f"{lane} baseline is not the original source fence")
    _expect(FULL_SHA.fullmatch(candidate) is not None and candidate != baseline, "candidate must be a distinct full source SHA")
    receipt_info, events = _validate_build_receipt(
        output_dir, lane, contract, baseline, candidate, before_target, after_target
    )
    reports: list[dict[str, Any]] = []
    for block in range(1, int(contract["blocks"]) + 1):
        for index, leg in enumerate(ABBA_ORDER, start=1):
            for program in contract["programs"]:
                stdout = events[(block, index, program)]["stdout"]
                if lane == "legacy-rtf":
                    report = _validate_rtf_log(stdout, contract["routes"])
                elif lane == "legacy-xls-properties":
                    report = _validate_metadata_log(stdout, contract["routes"])
                else:
                    route = {
                        "read": "read",
                        "edit": "edit",
                    }[program]
                    report = _validate_csv_log(stdout, route, int(contract["samples"]))
                report.update({"block": block, "leg": leg, "program": program})
                reports.append(report)
    return {
        "schema": SUMMARY_SCHEMA,
        "lane": lane,
        "baseline_source": baseline,
        "candidate_source": candidate,
        "abba_order": list(ABBA_ORDER),
        "abba_blocks": contract["blocks"],
        "routes": list(contract["routes"]),
        "probe_contract": contract["probe_contract"],
        "reports": reports,
        "source_lock_probe_binary_identity": receipt_info["identities"],
        "quiet_and_caps": {
            "process_guard": True,
            "target_preflight_bytes": TARGET_PRE_FLIGHT,
            "target_hard_bytes": TARGET_HARD,
            "memory_bytes": MEMORY_HARD,
        },
        "allocation": {
            "normal_timing_instrumentation": False,
            "separate_phase_required": True,
            "elapsed_claimable": False,
        },
        "performance_claim": False,
    }


def main(argv: list[str]) -> int:
    if len(argv) != 6:
        print(
            "usage: legacy_native_validator.py OUTPUT_DIR LANE BASELINE_SHA "
            "CANDIDATE_SHA BEFORE_TARGET AFTER_TARGET",
            file=sys.stderr,
        )
        return 2
    try:
        summary = validate_directory(
            Path(argv[0]),
            argv[1],
            argv[2],
            argv[3],
            Path(argv[4]),
            Path(argv[5]),
        )
        destination = Path(argv[0]) / "summary.json"
        _expect(not destination.exists() and not destination.is_symlink(), "summary.json already exists")
        with destination.open("x", encoding="utf-8") as stream:
            json.dump(summary, stream, indent=2, sort_keys=True)
            stream.write("\n")
        print(f"validated {len(summary['reports'])} legacy native reports")
    except (OSError, ValidationError) as error:
        print(f"legacy native qualification validation failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
