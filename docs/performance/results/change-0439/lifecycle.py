#!/usr/bin/env python3
"""Portable verifier and lifecycle gate for the 0439 ODP append baseline.

The capture and profile drivers own process execution.  This module only
checks retained receipts, source/binary/protocol custody, copied-oracle report
validation, profiles, sealing metadata, and cleanup proof.  It deliberately
has no before role and never authorizes a speedup or regression claim.

``--stage precleanup`` may return ``pending`` while capture is not complete;
that state is not a passing evidence gate.  ``--stage final`` requires the
complete 12-report/360-sample matrix, two profiles, and the sealed inventory.
The sidecar contract carries cleanup paths and final receipt names so this
verifier does not infer host paths from a previous batch.
"""

from __future__ import annotations

import argparse
import datetime as dt
import gzip
import hashlib
import importlib.util
import json
from pathlib import Path
import re
import subprocess
import sys
from typing import Any

ROOT = Path(__file__).resolve().parent
CHANGE = 439
STAGES = ("precleanup", "aftercleanup", "final")
MODES = ("normal", "allocator")
SHAPES = {"tiny": 64, "medium": 4_096, "large": 8_192}
PHASES = ("R1", "R2")
PROFILE_KINDS = ("stat", "record")
HEX64 = re.compile(r"^[0-9a-fA-F]{64}$")


class LifecycleError(ValueError):
    pass


class PendingEvidence(LifecycleError):
    def __init__(self, missing: list[str]):
        self.missing = missing
        super().__init__("evidence is not complete: " + ", ".join(missing))


def fail(label: str, message: str) -> None:
    raise LifecycleError(f"{label}: {message}")


def load_json(path: Path, label: str) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(label, f"invalid JSON: {error}")
    raise AssertionError("unreachable")


def obj(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(label, "expected an object")
    return value


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
    except OSError as error:
        fail(str(path), f"cannot hash file: {error}")
    return digest.hexdigest()


def hex_digest(value: Any, label: str) -> str:
    if not isinstance(value, str) or HEX64.fullmatch(value) is None:
        fail(label, "expected a SHA-256 digest")
    return value.lower()


def safe_relative(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        fail(label, "expected a non-empty bundle-relative path")
    path = Path(value)
    if path.is_absolute() or not path.parts or "." in path.parts or ".." in path.parts:
        fail(label, "absolute or traversal path is not allowed")
    return value


def bundle_path(value: Any, label: str) -> Path:
    relative = safe_relative(value, label)
    path = (ROOT / relative).resolve()
    if not path.is_relative_to(ROOT.resolve()):
        fail(label, "path escapes the evidence bundle")
    return path


def load_protocol() -> dict[str, Any]:
    path = ROOT / "protocol.json"
    value = obj(load_json(path, "protocol.json"), "protocol.json")
    if value.get("change") != CHANGE:
        fail("protocol.change", f"expected {CHANGE}")
    return value


def load_contract(protocol: dict[str, Any]) -> dict[str, Any]:
    path = ROOT / "lifecycle-contract.json"
    if not path.is_file():
        return {
            "schema": "litchi-0439-lifecycle-contract-v1",
            "change": CHANGE,
            "protocol_path": "protocol.json",
            "formal_receipts": None,
            "formal_profiles": None,
            "profile_driver": "profile.py",
            "required_checks": ["checks/after-build.json"],
            "cleanup": {},
            "drivers": ["capture.py", "profile.py", "profile-refined.py", "lifecycle.py", "seal.py", "replay.py", "portable-probes.py", "cleanup.py"],
        }
    value = obj(load_json(path, str(path)), "lifecycle-contract.json")
    if value.get("schema") != "litchi-0439-lifecycle-contract-v1" or value.get("change") != CHANGE:
        fail("lifecycle-contract.json", "schema or change differs")
    if value.get("protocol_path") != "protocol.json":
        fail("lifecycle-contract.json.protocol_path", "must be protocol.json")
    bound = value.get("protocol_sha256")
    if bound is not None and bound != sha(ROOT / "protocol.json"):
        fail("lifecycle-contract.json.protocol_sha256", "protocol binding is stale")
    return value


def expected_order(protocol: dict[str, Any]) -> list[tuple[str, str, str, str]]:
    expected = [
        ("R1", "normal", "tiny", "R1"),
        ("R1", "normal", "medium", "R1"),
        ("R1", "normal", "large", "R1"),
        ("R1", "allocator", "tiny", "R1"),
        ("R1", "allocator", "medium", "R1"),
        ("R1", "allocator", "large", "R1"),
        ("R2", "allocator", "large", "R2"),
        ("R2", "allocator", "medium", "R2"),
        ("R2", "allocator", "tiny", "R2"),
        ("R2", "normal", "large", "R2"),
        ("R2", "normal", "medium", "R2"),
        ("R2", "normal", "tiny", "R2"),
    ]
    rows = protocol.get("order")
    if not isinstance(rows, list) or len(rows) != len(expected):
        fail("protocol.order", "must contain exactly 12 lanes")
    actual: list[tuple[str, str, str, str]] = []
    for index, row in enumerate(rows):
        value = obj(row, f"protocol.order[{index}]")
        actual.append((
            value.get("phase"), value.get("mode"), value.get("shape"), value.get("repeat")
        ))
    if actual != expected:
        fail("protocol.order", "does not match the fixed R1/R2 matrix")
    return actual


def validate_protocol(protocol: dict[str, Any], stage: str, *, require_frozen: bool = False) -> list[tuple[str, str, str, str]]:
    if protocol.get("status") not in {"draft", "frozen"}:
        fail("protocol.status", "must be draft or frozen")
    if (require_frozen or stage in {"aftercleanup", "final"}) and protocol.get("status") != "frozen":
        fail("protocol.status", "final evidence requires a frozen protocol")
    if protocol.get("cpu") != 2 or protocol.get("workers") != 1:
        fail("protocol", "must bind CPU 2 and one worker")
    if protocol.get("samples") != 30 or protocol.get("warmups") != 3 or protocol.get("repeats") != 2:
        fail("protocol", "must bind 30 samples, 3 warmups, and two repeat phases")
    if protocol.get("modes") != list(MODES) or protocol.get("shapes") != SHAPES:
        fail("protocol", "mode and shape matrix differs from 0439 baseline")
    roles = obj(protocol.get("roles"), "protocol.roles")
    if set(roles) != {"after"}:
        fail("protocol.roles", "must contain only the after baseline role")
    role = obj(roles["after"], "protocol.roles.after")
    for key, expected in {
        "build_directory": "after",
        "selector": "odp_existing_append_lifecycle",
        "oracle_role": "after",
    }.items():
        if role.get(key) != expected:
            fail(f"protocol.roles.after.{key}", f"must be {expected!r}")
    if not isinstance(role.get("source_field"), str) or not role["source_field"]:
        fail("protocol.roles.after.source_field", "must bind a non-empty append fixture field")
    matrix = obj(protocol.get("matrix"), "protocol.matrix")
    if matrix.get("reports") != 12 or matrix.get("retained_samples") != 360:
        fail("protocol.matrix", "must bind 12 reports and 360 retained samples")
    if matrix.get("reports_per_phase") != 6 or matrix.get("samples_per_report") != 30:
        fail("protocol.matrix", "phase/report dimensions differ from 6 x 30")
    profile = obj(protocol.get("profiles"), "protocol.profiles")
    if profile.get("shape") != "large" or profile.get("mode") != "normal" or profile.get("kinds") != list(PROFILE_KINDS):
        fail("protocol.profiles", "must bind large normal stat and record profiles")
    claims = obj(protocol.get("claims"), "protocol.claims")
    withheld = claims.get("withheld")
    if not isinstance(withheld, list) or not any("before/after" in str(item) for item in withheld):
        fail("protocol.claims.withheld", "must explicitly withhold before/after claims")
    oracle = obj(protocol.get("oracle"), "protocol.oracle")
    verifier = oracle.get("verifier_path", oracle.get("path"))
    safe_relative(verifier, "protocol.oracle.verifier_path")
    return expected_order(protocol)


def validate_oracle_bindings(protocol: dict[str, Any]) -> tuple[Path, Path]:
    """Require both copied oracle files and the hashes frozen in protocol.json."""
    oracle = obj(protocol.get("oracle"), "protocol.oracle")
    verifier = bundle_path(oracle.get("verifier_path", oracle.get("path")), "protocol.oracle.verifier_path")
    verifier_sha = hex_digest(oracle.get("verifier_sha256"), "protocol.oracle.verifier_sha256")
    protocol_name = safe_relative(oracle.get("protocol_path"), "protocol.oracle.protocol_path")
    oracle_protocol = bundle_path(protocol_name, "protocol.oracle.protocol_path")
    oracle_protocol_sha = hex_digest(oracle.get("protocol_sha256"), "protocol.oracle.protocol_sha256")
    if not verifier.is_file() or sha(verifier) != verifier_sha:
        fail("protocol.oracle.verifier_sha256", "copied verifier is missing or differs from its frozen hash")
    if not oracle_protocol.is_file() or sha(oracle_protocol) != oracle_protocol_sha:
        fail("protocol.oracle.protocol_sha256", "copied oracle protocol is missing or differs from its frozen hash")
    return verifier, oracle_protocol


def optional_contract_paths(contract: dict[str, Any], key: str) -> tuple[str, ...] | None:
    value = contract.get(key)
    if value is None:
        return None
    if not isinstance(value, list) or any(not isinstance(item, str) for item in value):
        fail(f"lifecycle-contract.json.{key}", "must be a string path list")
    paths = tuple(safe_relative(item, f"lifecycle-contract.json.{key}") for item in value)
    if len(set(paths)) != len(paths):
        fail(f"lifecycle-contract.json.{key}", "contains duplicate paths")
    return paths


def artifact(path: Path, value: Any, label: str, *, required: bool = True) -> None:
    if not isinstance(value, dict):
        fail(label, "artifact record must be an object")
    member = bundle_path(value.get("path"), f"{label}.path")
    if not member.is_file():
        if required:
            compression_path = ROOT / "compression.json"
            if not compression_path.is_file():
                fail(label, "artifact file is missing")
            records = obj(load_json(compression_path, str(compression_path)), str(compression_path))
            original_name = value.get("path")
            matched = [
                (stored_name, item)
                for stored_name, item in records.items()
                if isinstance(item, dict) and item.get("original_path") == original_name
            ]
            if len(matched) != 1:
                fail(label, "artifact file is missing and has no unique compressed record")
            stored_name, item = matched[0]
            stored = bundle_path(stored_name, f"{label}.compressed_path")
            raw = gzip.decompress(stored.read_bytes())
            if len(raw) != value.get("bytes") or hashlib.sha256(raw).hexdigest() != value.get("sha256"):
                fail(label, "compressed artifact differs from receipt")
            return
        return
    size = value.get("bytes")
    if isinstance(size, bool) or not isinstance(size, int) or size < 0:
        fail(f"{label}.bytes", "must be a nonnegative integer")
    expected = hex_digest(value.get("sha256"), f"{label}.sha256")
    if member.stat().st_size != size or sha(member) != expected:
        fail(label, "artifact bytes or digest differ")


def source_manifest(value: Any, label: str) -> dict[str, Any]:
    row = obj(value, label)
    path_value = row.get("path")
    path = bundle_path(path_value, f"{label}.path")
    expected = hex_digest(row.get("sha256"), f"{label}.sha256")
    files = row.get("files")
    if isinstance(files, bool) or not isinstance(files, int) or files < 1:
        fail(f"{label}.files", "must be a positive integer")
    if not path.is_file() or path.stat().st_size == 0 or sha(path) != expected:
        fail(label, "retained source manifest is missing or stale")
    loaded = load_json(path, str(path))
    if not isinstance(loaded, dict) or len(loaded) != files:
        fail(label, "source manifest file count differs")
    return row


def validate_build(protocol_sha: str, stage: str) -> tuple[dict[str, Any] | None, list[str]]:
    path = ROOT / "after" / "build.json"
    if not path.is_file():
        return None, ["after/build.json"]
    row = obj(load_json(path, str(path)), str(path))
    if row.get("change") != CHANGE or row.get("role") not in {None, "after", "after-streaming"}:
        fail(str(path), "build descriptor identity differs")
    if row.get("protocol_sha256") != protocol_sha:
        fail(str(path), "build descriptor protocol binding is stale")
    source = source_manifest(row.get("source_manifest"), f"{path}.source_manifest")
    binaries = obj(row.get("binaries"), f"{path}.binaries")
    if set(binaries) != set(MODES):
        fail(str(path), "normal and allocator binary identities are required")
    for mode, identity in binaries.items():
        item = obj(identity, f"{path}.binaries.{mode}")
        size = item.get("bytes")
        if isinstance(size, bool) or not isinstance(size, int) or size <= 0:
            fail(f"{path}.binaries.{mode}.bytes", "must be positive")
        hex_digest(item.get("sha256"), f"{path}.binaries.{mode}.sha256")
        if not isinstance(item.get("path"), str) or not item["path"]:
            fail(f"{path}.binaries.{mode}.path", "is required")
    copies = ROOT / "after" / "binary-copies.json"
    if copies.is_file():
        copies_value = obj(load_json(copies, str(copies)), str(copies))
        if set(copies_value) != set(MODES):
            fail(str(copies), "normal and allocator copies are required")
        for mode in MODES:
            item = obj(copies_value[mode], f"{copies}.{mode}")
            if item.get("bytes") != binaries[mode].get("bytes") or item.get("sha256") != binaries[mode].get("sha256"):
                fail(str(copies), f"{mode} copy differs from build descriptor")
    return {"descriptor": row, "source": source}, []


def validate_source_receipt(path: Path, expected_status: str = "pass") -> None:
    row = obj(load_json(path, str(path)), str(path))
    if row.get("status") == "running":
        fail(str(path), "receipt is still running")
    if expected_status not in {"pass", "failed"}:
        fail(str(path), "expected check status must be pass or failed")
    if row.get("status") != expected_status:
        fail(str(path), f"source/check receipt status differs from expected {expected_status}")
    if expected_status == "pass" and row.get("exit_code") != 0:
        fail(str(path), "passing source/check receipt has a nonzero exit code")
    before = source_manifest(row.get("source_before"), f"{path}.source_before")
    after = source_manifest(row.get("source_after"), f"{path}.source_after")
    if before != after or row.get("source_unchanged") is not True:
        fail(str(path), "source custody was not unchanged")
    if isinstance(row.get("log"), dict):
        artifact(path, row["log"], f"{path}.log")


def validate_check_receipts(contract: dict[str, Any], stage: str) -> list[str]:
    required = optional_contract_paths(contract, "required_checks") or ("checks/after-build.json",)
    expected_value = contract.get("expected_checks", {})
    if not isinstance(expected_value, dict) or any(not isinstance(name, str) or not isinstance(status, str) for name, status in expected_value.items()):
        fail("lifecycle-contract.json.expected_checks", "must map bundle paths to terminal pass/failed statuses")
    expected = {safe_relative(name, "lifecycle-contract.json.expected_checks"): status for name, status in expected_value.items()}
    for name, status in expected.items():
        if status not in {"pass", "failed"}:
            fail(f"lifecycle-contract.json.expected_checks.{name}", "must be pass or failed")
    # Required checks are always passing gates, even if an accidental expected
    # map entry tries to classify one as a retained failed attempt.
    for name in required:
        if expected.get(name, "pass") != "pass":
            fail(f"lifecycle-contract.json.required_checks.{name}", "required checks must be expected to pass")
        expected[name] = "pass"
    names = tuple(dict.fromkeys((*required, *expected)))
    missing = [name for name in names if not bundle_path(name, "check receipt").is_file()]
    if missing:
        return missing
    for name in names:
        validate_source_receipt(bundle_path(name, name), expected[name])
    # Every retained source receipt not explicitly accounted for must itself
    # be passing. A failed attempt is acceptable only when the contract names
    # it in expected_checks with status "failed".
    if (ROOT / "checks").is_dir():
        for path in sorted((ROOT / "checks").glob("*.json")):
            value = load_json(path, str(path))
            if not isinstance(value, dict) or "source_before" not in value:
                continue
            name = str(path.relative_to(ROOT))
            if name in expected:
                continue
            if value.get("status") == "running":
                fail(name, "retained check is still running")
            if value.get("status") == "failed":
                fail(name, "failed retained check is missing from expected_checks")
            validate_source_receipt(path, "pass")
    return []


def oracle_command(protocol: dict[str, Any], report: Path, mode: str, shape: str) -> list[str]:
    oracle = obj(protocol["oracle"], "protocol.oracle")
    verifier = bundle_path(oracle.get("verifier_path", oracle.get("path")), "protocol.oracle.verifier_path")
    if not verifier.is_file():
        fail(str(verifier), "copied oracle verifier is missing")
    template = oracle.get("argv")
    if template is None:
        template = [sys.executable, "-B", "{verifier}", "--report", "{report}", "--mode", "{mode}", "--shape", "{shape}", "--role", "after"]
    if not isinstance(template, list) or any(not isinstance(item, str) for item in template):
        fail("protocol.oracle.argv", "must be a string list")
    values = {"python": sys.executable, "verifier": str(verifier), "report": str(report), "mode": mode, "shape": shape, "role": "after"}
    try:
        return [item.format(**values) for item in template]
    except (KeyError, ValueError) as error:
        fail("protocol.oracle.argv", f"invalid template: {error}")
    raise AssertionError("unreachable")


def report_build_identity(report: Path, build: dict[str, Any], mode: str) -> None:
    value = obj(load_json(report, str(report)), str(report))
    identity = obj(value.get("binary_identity"), "report.binary_identity")
    expected = build["binaries"][mode]
    if (identity.get("binary_sha256") != expected["sha256"]
            or identity.get("binary_bytes") != expected["bytes"]
            or identity.get("path") != expected["path"]
            or obj(value.get("environment"), "report.environment").get("git_revision") != build["revision"]):
        fail(str(report), "report executable/revision differs from retained build")


def run_oracle(protocol: dict[str, Any], report: Path, mode: str, shape: str, label: str) -> None:
    command = oracle_command(protocol, report, mode, shape)
    try:
        result = subprocess.run(command, cwd=ROOT, capture_output=True, text=True, check=False)
    except OSError as error:
        fail(label, f"oracle could not run: {error}")
    if result.returncode != 0 or result.stdout.strip() != "VALID":
        fail(label, "copied oracle rejected the report")


def artifact_is_nonempty(path: Path, value: Any, label: str) -> None:
    member = bundle_path(value.get("path"), f"{label}.path")
    if member.is_file():
        if member.stat().st_size == 0:
            fail(label, "required artifact is empty")
        return
    records_path = ROOT / "compression.json"
    if not records_path.is_file():
        fail(label, "required artifact is missing")
    records = obj(load_json(records_path, str(records_path)), str(records_path))
    matches = [item for item in records.values() if isinstance(item, dict) and item.get("original_path") == value.get("path")]
    if len(matches) != 1 or matches[0].get("original_bytes", 0) <= 0:
        fail(label, "required compressed artifact is empty or missing")


def receipt_artifacts(path: Path, row: dict[str, Any], required: tuple[str, ...], nonempty: tuple[str, ...] = ()) -> dict[str, Path]:
    values = obj(row.get("artifacts"), f"{path}.artifacts")
    result: dict[str, Path] = {}
    for key, value in values.items():
        artifact(path, value, f"{path}.artifacts.{key}")
        result[key] = bundle_path(value.get("path"), f"{path}.artifacts.{key}.path")
    for key in required:
        if key not in result:
            fail(str(path), f"required artifact {key!r} is missing")
    for key in nonempty:
        if key not in values:
            fail(str(path), f"required nonempty artifact {key!r} is missing")
        artifact_is_nonempty(path, values[key], f"{path}.artifacts.{key}")
    return result


def formal_receipt_paths(contract: dict[str, Any]) -> list[Path]:
    names = optional_contract_paths(contract, "formal_receipts")
    if names is not None:
        return [bundle_path(name, "formal receipt") for name in names]
    root = ROOT / "runs"
    if not root.is_dir():
        return []
    return sorted(
        path for path in root.glob("*/formal/*-receipt.json")
        if path.is_file()
    )


def receipt_times(row: dict[str, Any], path: Path, label: str) -> tuple[dt.datetime, dt.datetime]:
    try:
        started = dt.datetime.fromisoformat(str(row["started_utc"]))
        finished = dt.datetime.fromisoformat(str(row["finished_utc"]))
    except (KeyError, TypeError, ValueError) as error:
        fail(str(path), f"{label} timestamps are invalid: {error}")
    if started.tzinfo is None or finished.tzinfo is None or finished < started:
        fail(str(path), f"{label} timestamp order is invalid")
    return started, finished


def validate_capture_state(phase: str, receipt_paths: list[Path], protocol_sha: str) -> None:
    state = ROOT / "runs" / phase / "formal" / "capture-state.json"
    index = ROOT / "runs" / phase / "formal" / "capture-index.json"
    if not state.is_file() or not index.is_file():
        fail(f"runs/{phase}/formal", "capture state/index is missing")
    value = obj(load_json(state, str(state)), str(state))
    if value.get("status") != "pass" or value.get("phase") != phase or value.get("expected_lanes") != 6 or value.get("completed_lanes") != 6:
        fail(str(state), "capture state is not a passing six-lane state")
    if value.get("protocol_sha256") != protocol_sha:
        fail(str(state), "capture state protocol binding is stale")
    rows = load_json(index, str(index))
    if not isinstance(rows, list) or len(rows) != 6 or len(set(rows)) != 6:
        fail(str(index), "capture index must contain six unique receipt paths")
    expected = [str(path.relative_to(ROOT)) for path in receipt_paths]
    if rows != expected:
        fail(str(index), "capture index does not match the phase receipts")


def validate_formal_reports(protocol: dict[str, Any], contract: dict[str, Any], protocol_sha: str, build_info: dict[str, Any] | None) -> tuple[int, int, list[str], dt.datetime | None]:
    paths = formal_receipt_paths(contract)
    if not paths:
        return 0, 0, ["12 formal capture receipts"], None
    if len(paths) != 12:
        fail("runs", f"expected 12 formal receipts, found {len(paths)}")
    seen: set[tuple[str, str, str, str]] = set()
    by_key: dict[tuple[str, str, str, str], Path] = {}
    times: dict[tuple[str, str, str, str], tuple[dt.datetime, dt.datetime]] = {}
    expected_rows = expected_order(protocol)
    for path in paths:
        if not path.is_file():
            fail(str(path), "formal receipt is missing")
        row = obj(load_json(path, str(path)), str(path))
        if row.get("change") != CHANGE or row.get("status") != "pass" or row.get("exit_code") != 0 or row.get("oracle_exit_code") != 0:
            fail(str(path), "formal receipt is not a passing terminal receipt")
        if row.get("attempt") != "formal" or row.get("role") != "after":
            fail(str(path), "receipt is outside the formal after baseline")
        if row.get("protocol_sha256") != protocol_sha:
            fail(str(path), "receipt protocol binding is stale")
        if row.get("driver_sha256") != sha(ROOT / "capture.py"):
            fail(str(path), "capture driver custody is stale")
        lane = obj(row.get("lane"), f"{path}.lane")
        key = (lane.get("phase"), lane.get("mode"), lane.get("shape"), lane.get("repeat"))
        if key not in expected_rows or key in seen:
            fail(str(path), "lane is duplicated or outside the fixed matrix")
        seen.add(key)
        phase, mode, shape, repeat = key
        if row.get("phase") != phase or row.get("mode", mode) != mode or row.get("shape", shape) != shape:
            fail(str(path), "receipt lane fields disagree")
        if row.get("selector") != "odp_existing_append_lifecycle":
            fail(str(path), "selector binding differs")
        if row.get("source_unchanged") is not True or row.get("outside_bundle_status_unchanged") is not True:
            fail(str(path), "receipt does not prove source/repository custody")
        if row.get("source_before") != row.get("source_after"):
            fail(str(path), "source manifest changed during capture")
        times[key] = receipt_times(row, path, "capture")
        if build_info is not None:
            if row.get("source_before") != build_info["source"]:
                fail(str(path), "capture source manifest differs from after build")
            binary = obj(row.get("binary"), f"{path}.binary")
            expected_binary = obj(build_info["descriptor"]["binaries"].get(mode), f"after/build.json.binaries.{mode}")
            if binary.get("bytes") != expected_binary.get("bytes") or binary.get("sha256") != expected_binary.get("sha256"):
                fail(str(path), "capture binary identity differs from after build")
        artifacts = receipt_artifacts(path, row, ("report", "catalog"))
        if build_info is not None:
            report_build_identity(artifacts["report"], build_info["descriptor"], mode)
        run_oracle(protocol, artifacts["report"], mode, shape, str(path))
        by_key[key] = path
    if seen != set(expected_rows):
        fail("runs", "formal receipts do not cover exactly the 12 protocol lanes")
    capture_finished: dt.datetime | None = None
    for phase in PHASES:
        phase_rows = [key for key in expected_rows if key[0] == phase]
        if len(phase_rows) != 6:
            fail(f"runs/{phase}", "must contain exactly six formal lanes")
        ordered = [by_key[key] for key in phase_rows]
        validate_capture_state(phase, ordered, protocol_sha)
        previous_finished: dt.datetime | None = None
        for key in phase_rows:
            started, finished = times[key]
            if previous_finished is not None and started < previous_finished:
                fail(f"runs/{phase}", "capture lanes overlap or are out of order")
            previous_finished = finished
        if capture_finished is not None and times[phase_rows[0]][0] < capture_finished:
            fail("runs", "R2 capture begins before R1 capture has finished")
        capture_finished = times[phase_rows[-1]][1]
    return 12, 360, [], capture_finished


def profile_receipt_paths(contract: dict[str, Any]) -> list[Path]:
    names = optional_contract_paths(contract, "formal_profiles")
    if names is not None:
        return [bundle_path(name, "profile receipt") for name in names]
    return [ROOT / "profiles" / "after" / kind / "formal" / "receipt.json" for kind in PROFILE_KINDS]


def profile_driver_path(contract: dict[str, Any]) -> Path:
    name = safe_relative(contract.get("profile_driver"), "lifecycle-contract.profile_driver")
    path = bundle_path(name, "lifecycle-contract.profile_driver")
    if not path.is_file():
        fail("lifecycle-contract.profile_driver", "profile refinement driver is missing")
    return path


def cli_flag(protocol: dict[str, Any], name: str, default: str) -> str:
    cli = protocol.get("workload_cli", {})
    if not isinstance(cli, dict):
        fail("protocol.workload_cli", "must be an object")
    value = cli.get(name, default)
    if not isinstance(value, str) or not value:
        fail(f"protocol.workload_cli.{name}", "must be nonempty text")
    return value


def argv_path_matches(value: Any, artifact_path: Path) -> bool:
    if not isinstance(value, str) or not value:
        return False
    normalized = value.replace("\\", "/")
    relative = artifact_path.relative_to(ROOT).as_posix()
    return normalized == str(artifact_path) or normalized.endswith("/" + relative)


def expected_profile_argv(protocol: dict[str, Any], row: dict[str, Any], kind: str, artifacts: dict[str, Path]) -> list[Any]:
    profile = obj(protocol.get("profiles"), "protocol.profiles")
    role = obj(protocol.get("roles"), "protocol.roles")["after"]
    binary = obj(row.get("binary"), "profile.binary")
    binary_path = binary.get("path")
    if not isinstance(binary_path, str) or not binary_path:
        fail("profile.binary.path", "must be nonempty")
    base: list[Any] = [
        binary_path,
        cli_flag(protocol, "case_flag", "--case"), role.get("selector"),
        cli_flag(protocol, "shape_flag", "--semantic-shape"), "large",
        cli_flag(protocol, "workers_flag", "--workers"), str(protocol["workers"]),
        cli_flag(protocol, "samples_flag", "--samples"), str(protocol["samples"]),
        cli_flag(protocol, "warmup_flag", "--warmup"), str(protocol["warmups"]),
        cli_flag(protocol, "report_flag", "--json"), ("path", "report"),
        cli_flag(protocol, "catalog_flag", "--corpus-manifest"), ("path", "catalog"),
    ]
    resource = ("path", "resource")
    if kind == "stat":
        profiler: list[Any] = [
            "perf", "stat", "--no-big-num", "-x,", "-e", ",".join(profile.get("events", [])),
            "-o", ("path", "perf_stat"), "--", *base,
        ]
    else:
        profiler = [
            "perf", "record", "--no-buildid-cache", "-o", ("path", "perf_data"),
            "-F", str(profile.get("record_frequency_hz", 999)),
            "-e", str(profile.get("record_event", "cycles:u")),
            "--call-graph", str(profile.get("call_graph", "fp,127")), "--", *base,
        ]
    return ["taskset", "-c", str(protocol["cpu"]), "/usr/bin/time", "-v", "-o", resource, *profiler]


def assert_profile_argv(protocol: dict[str, Any], row: dict[str, Any], kind: str, artifacts: dict[str, Path]) -> None:
    actual = row.get("argv")
    expected = expected_profile_argv(protocol, row, kind, artifacts)
    if not isinstance(actual, list) or len(actual) != len(expected):
        fail("profile.argv", f"{kind} argv shape differs from the protocol-derived command")
    for index, (seen, wanted) in enumerate(zip(actual, expected)):
        if isinstance(wanted, tuple):
            if not argv_path_matches(seen, artifacts[wanted[1]]):
                fail("profile.argv", f"{kind} path binding differs at token {index}")
        elif seen != wanted:
            fail("profile.argv", f"{kind} token {index} differs from the protocol-derived command")


def profile_times(row: dict[str, Any], path: Path) -> tuple[dt.datetime, dt.datetime]:
    try:
        started = dt.datetime.fromisoformat(str(row["started_utc"]))
        finished = dt.datetime.fromisoformat(str(row["finished_utc"]))
    except (KeyError, TypeError, ValueError) as error:
        fail(str(path), f"profile timestamps are invalid: {error}")
    if started.tzinfo is None or finished.tzinfo is None or finished < started:
        fail(str(path), "profile timestamp order is invalid")
    return started, finished


def validate_profile_amendment(protocol_sha: str, build_info: dict[str, Any] | None) -> None:
    row = obj(load_json(ROOT / "profile-driver-amendment.json", "profile amendment"), "profile amendment")
    if row.get("change") != CHANGE:
        fail("profile amendment", "change differs")
    for key, name in (("original_driver", "profile.py"), ("refined_driver", "profile-refined.py")):
        binding = obj(row.get(key), key)
        if binding.get("path") != name or binding.get("sha256") != sha(ROOT / name):
            fail("profile amendment", f"{key} binding differs")
    if row.get("unchanged_capture_driver_sha256") != sha(ROOT / "capture.py") or row.get("unchanged_protocol_sha256") != protocol_sha:
        fail("profile amendment", "frozen capture/protocol binding differs")
    if build_info is not None and row.get("unchanged_source_manifest") != build_info["source"]:
        fail("profile amendment", "source binding differs")
    failed_name = "profiles/after/stat/formal/receipt.json"
    if row.get("failed_profile_receipt") != failed_name:
        fail("profile amendment", "failed attempt path differs")
    failed = obj(load_json(ROOT / failed_name, failed_name), failed_name)
    if failed.get("status") != "failed" or failed.get("driver_sha256") != sha(ROOT / "profile.py") or failed.get("protocol_sha256") != protocol_sha:
        fail("profile amendment", "original failed attempt binding differs")


def validate_profiles(protocol: dict[str, Any], contract: dict[str, Any], protocol_sha: str, build_info: dict[str, Any] | None, capture_finished: dt.datetime | None) -> tuple[int, list[str]]:
    validate_profile_amendment(protocol_sha, build_info)
    paths = profile_receipt_paths(contract)
    if not all(path.is_file() for path in paths):
        return 0, ["two formal profile receipts"]
    if len(paths) != 2:
        fail("profiles", "formal profile list must contain stat and record")
    expected_paths = [("stat", "recapture"), ("record", "formal")]
    actual_paths = []
    for path in paths:
        parts = path.relative_to(ROOT).parts
        if len(parts) == 5 and parts[0] == "profiles" and parts[1] == "after":
            actual_paths.append((parts[2], parts[3]))
        else:
            actual_paths.append((None, None))
    if actual_paths != expected_paths:
        fail("profiles", "profile paths must be stat/recapture followed by record/formal")
    profile_driver = profile_driver_path(contract)
    capture_helper = ROOT / "capture.py"
    if not capture_helper.is_file():
        fail("capture.py", "frozen capture helper is missing")
    capture_driver_sha = sha(capture_helper)
    profile_driver_sha = sha(profile_driver)
    seen: set[str] = set()
    profile = obj(protocol["profiles"], "protocol.profiles")
    events = profile.get("events")
    source_field = obj(protocol["roles"], "protocol.roles")["after"].get("source_field")
    if not isinstance(events, list) or any(not isinstance(item, str) or not item for item in events):
        fail("protocol.profiles.events", "must be a nonempty string list")
    previous_finished: dt.datetime | None = capture_finished
    for path in paths:
        parts = path.relative_to(ROOT).parts
        if len(parts) != 5 or parts[0] != "profiles" or parts[1] != "after" or parts[2] not in PROFILE_KINDS or parts[3] not in {"formal", "recapture"} or parts[4] != "receipt.json":
            fail(str(path), "profile path must be profiles/after/{stat/recapture,record/formal}/receipt.json")
        kind = parts[2]
        attempt = parts[3]
        row = obj(load_json(path, str(path)), str(path))
        if row.get("change") != CHANGE or row.get("status") != "pass" or row.get("exit_code") != 0 or row.get("oracle_exit_code") != 0 or row.get("role") != "after" or row.get("kind") != kind or row.get("attempt") != attempt:
            fail(str(path), "profile receipt is not a passing terminal profile")
        if row.get("protocol_sha256") != protocol_sha or row.get("driver_sha256") != profile_driver_sha:
            fail(str(path), "profile driver/protocol custody is stale")
        if row.get("capture_helper_sha256") != capture_driver_sha:
            fail(str(path), "profile capture-helper custody is stale")
        if row.get("status") != "pass" or row.get("source_unchanged") is not True or row.get("outside_bundle_status_unchanged") is not True or row.get("source_before") != row.get("source_after"):
            fail(str(path), "profile does not prove source/repository custody")
        if row.get("selector") != "odp_existing_append_lifecycle" or row.get("shape") != "large" or row.get("source_field") != source_field:
            fail(str(path), "profile selector/shape/source binding differs")
        if build_info is not None:
            if row.get("source_manifest") != build_info["source"]:
                fail(str(path), "profile source manifest differs from after build")
            binary = obj(row.get("binary"), f"{path}.binary")
            expected_binary = obj(build_info["descriptor"]["binaries"].get("normal"), "after/build.json.binaries.normal")
            if binary.get("bytes") != expected_binary.get("bytes") or binary.get("sha256") != expected_binary.get("sha256"):
                fail(str(path), "profile binary identity differs from after build")
        if kind == "stat" and row.get("stat_events") != events:
            fail(str(path), "perf stat event binding differs")
        if kind == "record" and (row.get("record_event") != profile.get("record_event", "cycles:u") or row.get("record_frequency_hz") != profile.get("record_frequency_hz", 999) or row.get("call_graph") != profile.get("call_graph", "fp,127")):
            fail(str(path), "perf record binding differs")
        started, finished = profile_times(row, path)
        if previous_finished is None or started < previous_finished:
            fail(str(path), "profile starts before the preceding capture/profile finished")
        previous_finished = finished
        required = ["report", "catalog", "resource", "workload_log", "oracle_log"]
        nonempty = ["report", "catalog", "resource", "oracle_log"]
        if kind == "stat":
            required.append("perf_stat")
            nonempty.append("perf_stat")
        else:
            required.extend(("perf_data", "perf_script", "perf_report"))
            nonempty.extend(("perf_data", "perf_script", "perf_report"))
        artifacts = receipt_artifacts(path, row, tuple(required), tuple(nonempty))
        assert_profile_argv(protocol, row, kind, artifacts)
        if build_info is not None:
            report_build_identity(artifacts["report"], build_info["descriptor"], "normal")
        run_oracle(protocol, artifacts["report"], "normal", "large", str(path))
        seen.add(kind)
    if seen != set(PROFILE_KINDS):
        fail("profiles", "stat and record profiles are both required")
    return 2, []


def inventory_rows(path: Path) -> dict[str, str]:
    if not path.is_file():
        fail(str(path), "sealed SHA256SUMS inventory is missing")
    rows: dict[str, str] = {}
    for line in path.read_text(encoding="utf-8").splitlines():
        if "  " not in line:
            fail(str(path), "inventory contains a malformed row")
        digest, name = line.split("  ", 1)
        if HEX64.fullmatch(digest) is None or not name or name in rows or name == "SHA256SUMS":
            fail(str(path), "inventory contains an unsafe or duplicate member")
        member = bundle_path(name, "inventory member")
        if not member.is_file() or sha(member) != digest.lower():
            fail(name, "inventory member is stale or missing")
        rows[name] = digest.lower()
    actual = {
        path.relative_to(ROOT).as_posix()
        for path in ROOT.rglob("*")
        if path.is_file() and path.name != "SHA256SUMS"
    }
    if actual != set(rows):
        fail("SHA256SUMS", "inventory does not cover the bundle exactly")
    return rows


def validate_compression() -> None:
    path = ROOT / "compression.json"
    if not path.is_file():
        return
    rows = obj(load_json(path, str(path)), str(path))
    import gzip
    for stored_name, record in rows.items():
        stored = bundle_path(stored_name, f"compression.{stored_name}")
        item = obj(record, f"compression.{stored_name}")
        raw = gzip.decompress(stored.read_bytes())
        if item.get("stored_bytes") != len(stored.read_bytes()) or item.get("stored_sha256") != sha(stored):
            fail(str(path), f"stored gzip metadata differs for {stored_name}")
        if item.get("original_bytes") != len(raw) or item.get("original_sha256") != hashlib.sha256(raw).hexdigest():
            fail(str(path), f"original gzip metadata differs for {stored_name}")
        original_name = item.get("original_path")
        if isinstance(original_name, str) and (ROOT / original_name).is_file():
            original = bundle_path(original_name, f"compression.{stored_name}.original_path")
            if original.read_bytes() != raw:
                fail(str(path), f"retained original differs for {stored_name}")


def cleanup_status(contract: dict[str, Any], stage: str, *, local_check: bool = False) -> dict[str, Any]:
    cleanup = obj(contract.get("cleanup", {}), "lifecycle-contract.cleanup")
    if stage == "precleanup":
        return {"status": "not-required", "stage": stage}
    receipt_name = cleanup.get("precleanup_receipt")
    inventory_name = cleanup.get("cleanup_inventory")
    paths = cleanup.get("paths")
    preserved = cleanup.get("preserved_paths")
    if not isinstance(receipt_name, str) or not isinstance(inventory_name, str) or not isinstance(paths, list) or not isinstance(preserved, list):
        fail("lifecycle-contract.cleanup", "final cleanup bindings are incomplete")
    receipt = bundle_path(receipt_name, "cleanup.precleanup_receipt")
    inventory = bundle_path(inventory_name, "cleanup.cleanup_inventory")
    if not receipt.is_file() or not inventory.is_file():
        fail("cleanup", "cleanup receipt/inventory is missing")
    row = obj(load_json(receipt, str(receipt)), str(receipt))
    if row.get("status") != "pass":
        fail(str(receipt), "precleanup lifecycle receipt is not passing")
    cleanup_row = obj(load_json(inventory, str(inventory)), str(inventory))
    if cleanup_row.get("status") != "pass" or cleanup_row.get("change") != CHANGE:
        fail(str(inventory), "cleanup inventory is not passing")
    prefix = cleanup.get("prefix", "litchi-goal-0439-")
    if cleanup_row.get("cleanup_paths") != paths:
        fail(str(inventory), "cleanup inventory paths differ from the contract")
    identities = cleanup_row.get("preserved_target_directory_identity")
    if not isinstance(identities, dict) or set(identities) != set(str(item) for item in preserved):
        fail(str(inventory), "cleanup inventory preserved-target set differs")
    for raw in paths:
        target = Path(raw)
        if not target.is_absolute() or target.parent != Path("/tmp") or not target.name.startswith(prefix):
            fail("cleanup.paths", f"unsafe cleanup path {raw!r}")
        if local_check and target.exists():
            fail("cleanup", f"owned cleanup path remains: {target}")
    for raw in preserved:
        target = Path(raw)
        if not target.is_absolute() or ".." in target.parts or not target.name:
            fail("cleanup.preserved_paths", f"unsafe preserved path {raw!r}")
    goal = cleanup.get("goal_path")
    goal_sha = cleanup.get("goal_sha256")
    if not isinstance(goal, str) or not Path(goal).is_absolute() or Path(goal).name != "GOAL.md":
        fail("cleanup.goal", "goal_path must be an absolute GOAL.md binding")
    goal_sha = hex_digest(goal_sha, "cleanup.goal_sha256")
    if cleanup_row.get("goal_path") != goal or cleanup_row.get("goal_sha256") != goal_sha:
        fail(str(inventory), "cleanup inventory GOAL binding differs")
    if local_check:
        goal_path = Path(goal)
        if not goal_path.is_file() or goal_path.is_symlink() or sha(goal_path) != goal_sha:
            fail("cleanup.goal", "preserved GOAL.md identity differs")
    return {"status": "pass", "stage": stage, "removed": len(paths), "preserved": len(preserved)}


def validate(stage: str, *, allow_pending: bool = False, require_inventory: bool = True, require_frozen: bool = False, local_cleanup: bool = False) -> dict[str, Any]:
    protocol = load_protocol()
    contract = load_contract(protocol)
    order = validate_protocol(protocol, stage, require_frozen=require_frozen)
    protocol_sha = sha(ROOT / "protocol.json")
    validate_oracle_bindings(protocol)
    build, missing_build = validate_build(protocol_sha, stage)
    missing: list[str] = list(missing_build)
    checks_missing = validate_check_receipts(contract, stage)
    missing.extend(checks_missing)
    reports, samples, missing_reports, capture_finished = validate_formal_reports(protocol, contract, protocol_sha, build)
    missing.extend(missing_reports)
    profiles, missing_profiles = validate_profiles(protocol, contract, protocol_sha, build, capture_finished)
    missing.extend(missing_profiles)
    if require_inventory and not (ROOT / "SHA256SUMS").is_file():
        missing.append("SHA256SUMS")
    if ROOT.joinpath("SHA256SUMS").is_file() and require_inventory:
        inventory_rows(ROOT / "SHA256SUMS")
        validate_compression()
    cleanup = cleanup_status(contract, stage, local_check=local_cleanup) if stage != "precleanup" else {"status": "not-required", "stage": stage}
    if missing:
        # A pre-capture draft is explicitly observable as pending.  Once any
        # terminal receipt exists, incomplete/failed matrices still fail.
        terminal_present = reports > 0 or profiles > 0 or build is not None
        if allow_pending and stage == "precleanup" and not terminal_present:
            return {"status": "pending", "change": CHANGE, "stage": stage, "missing": sorted(set(missing)), "reports": reports, "samples": samples, "profiles": profiles}
        raise PendingEvidence(sorted(set(missing)))
    if reports != 12 or samples != 360 or profiles != 2:
        fail("matrix", "complete evidence dimensions were not retained")
    return {
        "status": "pass",
        "change": CHANGE,
        "stage": stage,
        "reports": reports,
        "retained_samples": samples,
        "profiles": profiles,
        "before_after_claim": False,
        "cleanup": cleanup,
        "protocol_sha256": protocol_sha,
        "order": [list(item) for item in order],
    }


def load_probe_module():
    path = ROOT / "portable-probes.py"
    if not path.is_file():
        fail("portable-probes", "copied probe driver is missing")
    spec = importlib.util.spec_from_file_location("change0439_portable_probes", path)
    if spec is None or spec.loader is None:
        fail("portable-probes", "cannot load copied probe driver")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--stage", choices=STAGES, default="precleanup")
    parser.add_argument("--allow-pending", action="store_true")
    parser.add_argument("--seal", action="store_true", help="validate complete evidence, then invoke deterministic seal")
    parser.add_argument("--portable-probes", action="store_true", help="run copied-bundle mutation probes after a final gate")
    parser.add_argument("--local-cleanup-check", action="store_true", help="also inspect live absolute cleanup/GOAL paths; omitted for portable replay")
    args = parser.parse_args(argv)
    try:
        if args.seal:
            result = validate(args.stage, allow_pending=False, require_inventory=False, require_frozen=True, local_cleanup=args.local_cleanup_check)
            from seal import seal_bundle
            result["seal"] = seal_bundle()
        else:
            result = validate(args.stage, allow_pending=args.allow_pending, require_inventory=args.stage != "precleanup", local_cleanup=args.local_cleanup_check)
        if args.portable_probes:
            if result.get("status") != "pass":
                fail("portable-probes", "requires a passing complete lifecycle")
            result["portable_probes"] = load_probe_module().run_probes(args.stage)
        print(json.dumps(result, indent=2, sort_keys=True))
        return 0
    except PendingEvidence as error:
        print(json.dumps({"status": "pending", "change": CHANGE, "stage": args.stage, "missing": error.missing}, sort_keys=True))
        return 1
    except (OSError, KeyError, TypeError, ValueError, AssertionError, LifecycleError) as error:
        print(f"INVALID: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
