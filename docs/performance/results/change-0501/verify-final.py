#!/usr/bin/env python3
"""Independently verify the complete 0501 evidence bundle.

This verifier deliberately does not import or execute compare.py.  It validates
historical before evidence through its retained source archives, validates the
candidate against the current after freeze, checks every retained report and
phase counter, and recomputes the comparison object in memory.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import math
import re
import statistics
import sys
from pathlib import Path
from typing import Any, Iterable


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
TMP_ROOT = Path("/tmp/litchi-goal-0501").resolve()
CHANGE = 501
HEX40 = re.compile(r"^[0-9a-f]{40}$")
HEX64 = re.compile(r"^[0-9a-f]{64}$")
TIMINGS = (
    "api_sum_ns",
    "open_ns",
    "plan_ns",
    "publication_ns",
    "open_source_ns",
    "open_destination_ns",
)
READ_COUNTERS = ("logical_calls", "requested_bytes", "returned_bytes", "short_reads", "delayed_calls")
CACHE_COUNTERS = (
    "hits",
    "cold_loads",
    "waiter_joins",
    "successful_loads",
    "failed_loads",
    "evictions",
    "bypasses",
    "oversized_bypasses",
    "allocation_bypasses",
    "budget_reservation_failures",
)
BUDGET_COUNTERS = (
    "memory_used",
    "input_bytes_used",
    "output_bytes_used",
    "work_used",
    "objects_used",
    "depth_used",
)
CORE_ARTIFACTS = ("report", "resource", "stdout", "oracle")
PROFILE_ARTIFACTS = (
    "stat",
    "report",
    "resource_stat",
    "resource_record",
    "stat_stdout",
    "record_stdout",
    "report_json",
    "record_report",
)
PROFILE_SUFFIXES = {
    "stat": ".perf-stat.txt",
    "report": ".perf-report.txt",
    "resource_stat": ".stat.time.txt",
    "resource_record": ".record.time.txt",
    "stat_stdout": ".stat.stdout.txt",
    "record_stdout": ".record.stdout.txt",
    "report_json": ".report.json",
    "record_report": ".record.report.json",
}
GATE_CPU_PREFIX = ("taskset", "-c", "16-31")


class CheckError(ValueError):
    """Evidence did not satisfy the frozen protocol."""


def fail(message: str) -> None:
    raise CheckError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def sha_bytes(raw: bytes) -> str:
    return hashlib.sha256(raw).hexdigest()


def sha_file(path: Path) -> str:
    if not path.is_file():
        fail(f"missing file: {path}")
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def canonical_bytes(value: Any) -> bytes:
    return (json.dumps(value, sort_keys=True, indent=2) + "\n").encode()


def load_json(path: Path) -> Any:
    def duplicate(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            if key in result:
                fail(f"{path}: duplicate JSON key {key}")
            result[key] = value
        return result

    def nonfinite(value: str) -> Any:
        fail(f"{path}: non-finite JSON constant {value}")

    if not path.is_file():
        fail(f"missing JSON file: {path}")
    try:
        value = json.loads(path.read_text(), object_pairs_hook=duplicate, parse_constant=nonfinite)
    except CheckError:
        raise
    except (OSError, json.JSONDecodeError) as error:
        fail(f"{path}: invalid JSON: {error}")
    return value


def safe_bundle(relative: str) -> Path:
    candidate = Path(relative)
    if candidate.is_absolute() or ".." in candidate.parts:
        fail(f"unsafe bundle path: {relative}")
    resolved = (HERE / candidate).resolve()
    if not resolved.is_relative_to(HERE):
        fail(f"bundle path escapes result directory: {relative}")
    return resolved


def safe_repo(relative: str) -> Path:
    candidate = Path(relative)
    if candidate.is_absolute() or ".." in candidate.parts:
        fail(f"unsafe repository path: {relative}")
    resolved = (REPO / candidate).resolve()
    if not resolved.is_relative_to(REPO):
        fail(f"repository path escapes repository: {relative}")
    return resolved


def text(value: Any, label: str) -> str:
    require(isinstance(value, str) and bool(value), f"{label}: expected non-empty string")
    return value


def digest(value: Any, label: str) -> str:
    value = text(value, label)
    require(bool(HEX64.fullmatch(value)), f"{label}: expected lowercase SHA-256")
    return value


def integer(value: Any, label: str) -> int:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0, f"{label}: expected unsigned integer")
    return value


def module_from(path: Path, name: str) -> Any:
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        fail(f"cannot import {path}")
    module = importlib.util.module_from_spec(spec)
    try:
        spec.loader.exec_module(module)
    except Exception as error:
        fail(f"cannot import {path}: {error}")
    return module


def checked_report(verifier: Any, report: dict[str, Any], samples: int, warmups: int) -> dict[str, Any]:
    try:
        return verifier.check_report(report, expected_samples=samples, expected_warmup=warmups)
    except TypeError as error:
        # The historical formal helper predates the explicit supplementary
        # dimensions.  It is safe to use only for the frozen 30/3 core.
        require(samples == 30 and warmups == 3 and "unexpected keyword argument" in str(error), f"report checker call failed: {error}")
        return verifier.check_report(report)


def bound_path(name: str) -> Path:
    if name in {
        "protocol.json",
        "custody.py",
        "capture.py",
        "verify-report.py",
        "verify.py",
        "compare.py",
        "profile.py",
        "freeze.py",
    }:
        return HERE / name
    return safe_repo(name)


def source_files(value: Any, label: str) -> dict[str, str]:
    if isinstance(value, dict) and isinstance(value.get("source_manifest"), dict):
        value = value["source_manifest"]
    require(isinstance(value, dict), f"{label}: expected object")
    files = value.get("files")
    require(isinstance(files, dict) and files, f"{label}.files: expected non-empty object")
    result: dict[str, str] = {}
    for name, value in files.items():
        text(name, f"{label}.files path")
        digest(value, f"{label}.files[{name!r}]")
        safe_repo(name)
        result[name] = value
    return result


def validate_source_identity(manifest: dict[str, Any], label: str) -> dict[str, str]:
    files = source_files(manifest, label)
    require(manifest.get("files_count") == len(files), f"{label}.files_count mismatch")
    require(manifest.get("sha256") == sha_bytes(canonical_bytes(files)), f"{label}.sha256 mismatch")
    return files


def selected_source_files(path: Path) -> tuple[str | None, dict[str, str]]:
    data = load_json(path)
    revision = data.get("revision") if isinstance(data, dict) else None
    return revision, source_files(data, str(path))


def verify_protected_work() -> dict[str, Any]:
    data = load_json(HERE / "protected-work.json")
    require(isinstance(data, list) and data, "protected-work.json: expected non-empty list")
    rows = []
    for index, row in enumerate(data):
        require(isinstance(row, dict), f"protected-work[{index}]: expected object")
        name = text(row.get("path"), f"protected-work[{index}].path")
        expected = digest(row.get("sha256"), f"protected-work[{index}].sha256")
        path = safe_repo(name)
        actual = sha_file(path)
        require(actual == expected, f"protected path changed: {name}")
        rows.append({"path": name, "sha256": actual})
    return {"files": len(rows), "sha256": sha_bytes(canonical_bytes(rows)), "rows": rows}


def locate_old_helper(path_name: str, expected_sha: str) -> Path | None:
    candidates = []
    if path_name == "verify-report.py":
        candidates.append(HERE / "interrupted/profile-attempt-2/verify-report-before-dimension-fix.py")
    candidates.extend(
        [
            HERE / "interrupted" / f"{path_name}.before",
            HERE / "interrupted" / f"{path_name}.before-freeze",
            HERE / "interrupted" / f"{path_name[:-3]}-before.py",
            HERE / "interrupted" / f"{path_name[:-3]}-before-freeze.py",
            HERE / "source-archives/before" / f"{path_name}.txt",
            HERE / "source-archives/before" / f"{path_name}",
        ]
    )
    for path in candidates:
        if path.is_file() and sha_file(path) == expected_sha:
            return path

    # The coordinator may record custody in a small JSON manifest.  Accept
    # only an entry that names the exact historical digest and retained path.
    json_candidates = list(HERE.glob("*custody*.json")) + list(HERE.glob("*transition*.json"))
    json_candidates += list((HERE / "source-archives").glob("*.json"))
    seen: set[Path] = set()
    for manifest_path in json_candidates:
        if manifest_path in seen or not manifest_path.is_file():
            continue
        seen.add(manifest_path)
        try:
            manifest = load_json(manifest_path)
        except CheckError:
            continue
        for node in walk_dicts(manifest):
            if not isinstance(node, dict):
                continue
            names = [str(key) for key in node]
            if not any(path_name in name or path_name.removesuffix(".py") in name for name in names):
                continue
            found_sha = None
            found_path = None
            for key, value in node.items():
                key_lower = str(key).lower()
                if key_lower in {"sha256", "before_sha256", "source_sha256", "historical_sha256"} and isinstance(value, str):
                    found_sha = value
                if key_lower in {"path", "archive", "source", "before", "historical_path", "retained_path"} and isinstance(value, str):
                    found_path = value
            if found_sha != expected_sha or found_path is None:
                continue
            path = safe_bundle(found_path) if not Path(found_path).is_absolute() else Path(found_path).resolve()
            if path.is_file() and sha_file(path) == expected_sha:
                return path
    return None


def verify_driver_provenance(before_freeze: dict[str, Any], *, strict: bool) -> dict[str, Any]:
    transition_path = HERE / "driver-transition.json"
    transition = load_json(transition_path)
    require(transition.get("change") == CHANGE, "driver-transition.json: change mismatch")
    before_hashes = transition.get("bound_driver_hashes_at_before_freeze")
    current_hashes = transition.get("current_driver_hashes")
    require(isinstance(before_hashes, dict) and isinstance(current_hashes, dict), "driver-transition: missing driver hash maps")
    for name, expected in before_hashes.items():
        require(before_freeze["bound_files"].get(name) == expected, f"driver-transition before hash mismatch: {name}")
    for name, expected in current_hashes.items():
        require(digest(expected, f"driver-transition current {name}") == sha_file(bound_path(name)), f"current driver hash mismatch: {name}")
    changes = transition.get("driver_changes")
    require(isinstance(changes, list), "driver-transition: missing change records")
    custody: dict[str, Any] = {}
    for name, expected in before_hashes.items():
        old = locate_old_helper(name, expected)
        if old is None:
            if strict:
                fail(f"historical {name} custody is missing for before verifier")
            custody[name] = {"status": "missing", "sha256": expected}
        else:
            custody[name] = {"status": "verified", "path": str(old.relative_to(HERE) if old.is_relative_to(HERE) else old), "sha256": sha_file(old)}
    return {"manifest": str(transition_path.relative_to(HERE)), "changes": changes, "historical_helpers": custody}


def verify_freeze(
    phase: str,
    protocol: dict[str, Any],
    before_freeze: dict[str, Any] | None,
    *,
    strict_driver_custody: bool,
) -> tuple[dict[str, Any], dict[str, str], dict[str, Any]]:
    freeze_path = HERE / f"{phase}-freeze.json"
    freeze = load_json(freeze_path)
    require(freeze.get("change") == CHANGE and freeze.get("phase") == phase, f"{phase}-freeze identity mismatch")
    protocol_hash = sha_file(HERE / "protocol.json")
    require(freeze.get("protocol_sha256") == protocol_hash, f"{phase}: protocol changed after freeze")
    require(freeze.get("harness") == protocol["harness"], f"{phase}: harness differs from protocol")
    harness_path = safe_repo(protocol["harness"])
    require(sha_file(harness_path) == freeze["bound_files"].get(protocol["harness"]), f"{phase}: current harness differs from freeze")
    binary = freeze.get("binary")
    require(isinstance(binary, dict), f"{phase}: missing binary receipt")
    binary_path = Path(text(binary.get("path"), f"{phase}.binary.path")).resolve()
    require(binary_path.is_file(), f"{phase}: frozen binary is missing")
    require(integer(binary.get("bytes"), f"{phase}.binary.bytes") == binary_path.stat().st_size, f"{phase}: binary size changed")
    require(digest(binary.get("sha256"), f"{phase}.binary.sha256") == sha_file(binary_path), f"{phase}: frozen binary hash changed")
    source_manifest = freeze.get("source_manifest")
    require(isinstance(source_manifest, dict), f"{phase}: missing source manifest")
    files = validate_source_identity(source_manifest, f"{phase}.source_manifest")
    driver = {}
    if phase == "before":
        require(isinstance(freeze.get("bound_files"), dict), f"{phase}: missing bound file hashes")
        transition = load_json(HERE / "driver-transition.json")
        transition_current = transition.get("current_driver_hashes", {})
        allowed_historical_changes = {"crates/litchi-pptx/src/presentation/source_cross_copy.rs", "profile.py", "verify-report.py"}
        for name, expected in freeze["bound_files"].items():
            digest(expected, f"{phase}.bound_files[{name!r}]")
            actual = sha_file(bound_path(name))
            if name in allowed_historical_changes:
                if name in transition_current:
                    require(actual == transition_current[name], f"{phase}: changed helper does not match driver-transition: {name}")
                else:
                    require(name == "crates/litchi-pptx/src/presentation/source_cross_copy.rs", f"{phase}: undocumented historical bound-file change: {name}")
            else:
                require(actual == expected, f"{phase}: bound file changed: {name}")
        driver = verify_driver_provenance(freeze, strict=strict_driver_custody)
        custody = load_json(HERE / "source-custody-before.json")
        require(custody.get("revision") == freeze["revision"], "source-custody-before revision mismatch")
        custody_files = source_files(custody, "source-custody-before")
        require(set(custody_files) <= set(files), "source-custody-before contains files outside before freeze")
        for name, expected in custody_files.items():
            require(files[name] == expected, f"source-custody-before hash differs from before freeze: {name}")
        archives = load_json(HERE / "source-archives/before.json")
        require(isinstance(archives, dict) and archives, "source-archives/before.json: empty")
        archived_names = set(archives)
        for name, metadata in archives.items():
            require(name in files, f"before archive is outside historical manifest: {name}")
            require(isinstance(metadata, dict), f"before archive metadata: {name}")
            archive_path = safe_bundle(text(metadata.get("archive"), f"archive[{name}].archive"))
            expected = digest(metadata.get("sha256"), f"archive[{name}].sha256")
            expected_bytes = integer(metadata.get("bytes"), f"archive[{name}].bytes")
            require(expected == files[name], f"before archive hash differs from manifest: {name}")
            require(archive_path.is_file(), f"missing historical source archive: {archive_path}")
            require(archive_path.stat().st_size == expected_bytes, f"historical archive size mismatch: {name}")
            require(sha_file(archive_path) == expected, f"historical archive hash mismatch: {name}")
        selected_revision, selected = selected_source_files(HERE / "before-source.json")
        require(selected_revision == freeze["revision"], "before-source revision mismatch")
        for name, expected in selected.items():
            require(name in files and files[name] == expected, f"before-source hash differs from freeze: {name}")
        # Every historical file that has no retained archive must still be
        # present and unchanged.  Archived paths are intentionally allowed to
        # differ because they are the edited production/test files.
        for name, expected in files.items():
            if name in archived_names:
                continue
            require(sha_file(safe_repo(name)) == expected, f"unarchived historical source changed: {name}")
    else:
        require(isinstance(freeze.get("bound_files"), dict), f"{phase}: missing bound file hashes")
        for name, expected in freeze["bound_files"].items():
            digest(expected, f"{phase}.bound_files[{name!r}]")
            require(sha_file(bound_path(name)) == expected, f"{phase}: bound file changed: {name}")
        after_source_path = HERE / "after-source.json"
        selected_revision, selected = selected_source_files(after_source_path)
        require(selected_revision is None or selected_revision == freeze["revision"], "after-source revision mismatch")
        require(protocol["candidate"] in selected, "after-source does not bind the candidate source")
        require(protocol["harness"] in selected, "after-source does not bind the harness")
        for name, expected in selected.items():
            require(sha_file(safe_repo(name)) == expected, f"after-source hash mismatch: {name}")
        # The after freeze itself is a full current source identity.  This
        # protects every source file even if after-source.json is a selected
        # custody map.
        for name, expected in files.items():
            require(sha_file(safe_repo(name)) == expected, f"after source differs from after freeze: {name}")
        after_archives = load_json(HERE / "source-archives/after.json")
        require(isinstance(after_archives, dict) and after_archives, "source-archives/after.json: empty")
        for name, metadata in after_archives.items():
            require(name in files, f"after archive is outside after freeze: {name}")
            require(isinstance(metadata, dict), f"after archive metadata: {name}")
            archive_path = safe_bundle(text(metadata.get("archive"), f"after archive[{name}].archive"))
            expected = digest(metadata.get("sha256"), f"after archive[{name}].sha256")
            expected_bytes = integer(metadata.get("bytes"), f"after archive[{name}].bytes")
            require(expected == files[name], f"after archive hash differs from after freeze: {name}")
            require(archive_path.is_file() and archive_path.stat().st_size == expected_bytes, f"after archive size mismatch: {name}")
            require(sha_file(archive_path) == expected, f"after archive hash mismatch: {name}")

    return freeze, files, driver


def verify_frozen_relationship(before: dict[str, Any], after: dict[str, Any]) -> None:
    require(before["protocol_sha256"] == after["protocol_sha256"], "before/after protocol hashes differ")
    require(before["harness"] == after["harness"], "before/after harness paths differ")
    harness = before["harness"]
    require(before["bound_files"].get(harness) == after["bound_files"].get(harness), "before/after harness source differs")
    require(sha_file(safe_repo(harness)) == before["bound_files"][harness], "current harness is not the identical frozen harness")
    for name in ("capture.py", "compare.py", "custody.py", "freeze.py", "verify.py", "protocol.json"):
        require(before["bound_files"].get(name) == after["bound_files"].get(name), f"driver {name} differs between phases")


def verify_build(phase: str, freeze: dict[str, Any], source_receipt_path: Path) -> dict[str, Any]:
    if phase == "after":
        # The coordinator retains a small source-bound build receipt derived
        # from the benchmark-build gate.  Validate both records below.
        path = HERE / "build-after.json"
        name = path.name
    else:
        name = "build-before.json"
        path = HERE / name
    build = load_json(path)
    require(build.get("exit_code") == 0, f"{name}: build did not pass")
    expected_command = [
        "cargo",
        "build",
        "--release",
        "--locked",
        "--offline",
        "--manifest-path",
        "tools/perf-baseline/Cargo.toml",
        "--bin",
        "litchi-perf-baseline",
    ]
    require(build.get("command") == expected_command, f"{name}: build command differs")
    if phase == "before":
        require(build.get("rust_toolchain") == "1.98.1", f"{name}: toolchain differs")
        require(build.get("profile") == "release" and build.get("debug") == 0 and build.get("incremental") is False, f"{name}: build flags differ")
        require(build.get("target") == "/tmp/litchi-goal-0501-target", f"{name}: target differs")
        source_hash = sha_file(source_receipt_path)
    else:
        source_hash = sha_file(source_receipt_path)
        gate_receipt = HERE / "checks/benchmark-build.json"
        gate = load_json(gate_receipt)
        require(build.get("derived_from_gate") == "checks/benchmark-build.json", f"{name}: missing gate provenance")
        require(build.get("gate_receipt_sha256") == sha_file(gate_receipt), f"{name}: gate receipt hash mismatch")
        require(gate.get("exit_code") == 0 and gate.get("command") == list(GATE_CPU_PREFIX) + expected_command, f"{name}: benchmark-build gate differs")
    log = HERE / ("build-before.log" if phase == "before" else "build-after.log")
    require(digest(build.get("log_sha256"), f"{name}.log_sha256") == sha_file(log), f"{name}: build log hash mismatch")
    require(build.get("source_receipt_sha256") == source_hash, f"{name}: source custody receipt hash mismatch")
    require(build.get("retained_binary") == freeze["binary"]["path"], f"{name}: retained binary path differs")
    require(build.get("binary_sha256") == freeze["binary"]["sha256"], f"{name}: retained binary hash differs")
    return {
        "receipt": str(path.relative_to(HERE)),
        "log_sha256": build["log_sha256"],
        "source_receipt_sha256": build["source_receipt_sha256"],
        "binary_sha256": build["binary_sha256"],
    }


def resource_rss(path: Path) -> int:
    matches = re.findall(
        r"^\s*Maximum resident set size \(kbytes\):\s*(\d+)\s*$",
        path.read_text(errors="replace"),
        re.MULTILINE,
    )
    require(len(matches) == 1, f"{path}: expected exactly one GNU-time RSS line")
    return int(matches[0]) * 1024


def report_identity(report: dict[str, Any]) -> dict[str, Any]:
    return {
        "source_archive_sha256": report["source_archive_sha256"],
        "destination_archive_sha256": report["destination_archive_sha256"],
        "expected_output_sha256": report["expected_output_sha256"],
        "expected_output_bytes": report["expected_output_bytes"],
        "source_revision": report["source_revision"],
    }


def output_identity(report: dict[str, Any]) -> tuple[Any, ...]:
    return (
        report["source_archive_sha256"],
        report["source_archive_bytes"],
        report["destination_archive_sha256"],
        report["destination_archive_bytes"],
        report["expected_output_sha256"],
        report["expected_output_bytes"],
    )


def counter_fingerprint(report: dict[str, Any]) -> list[Any]:
    result = []
    for row in report["samples_raw"]:
        sample = []
        for phase in row["phases"]:
            owners = []
            for owner in ("source", "destination"):
                read = phase[f"{owner}_reads"]
                if read["availability"] == "unavailable":
                    read_value = None
                else:
                    read_value = tuple(read[key] for key in READ_COUNTERS) + tuple(read["request_size_counts"])
                cache = phase[f"{owner}_cache"]
                if cache["availability"] == "unavailable":
                    cache_value = None
                else:
                    cache_value = tuple(cache[key] for key in CACHE_COUNTERS)
                budget = phase[f"{owner}_budget"]
                budget_value = tuple(budget[key] for key in BUDGET_COUNTERS)
                owners.append((read_value, cache_value, budget_value))
            sample.append((phase["label"], tuple(owners)))
        result.append(tuple(sample))
    return result


def validate_artifacts(
    receipt: dict[str, Any],
    phase: str,
    name: str,
    expected_keys: Iterable[str],
    suffixes: dict[str, str],
) -> dict[str, Path]:
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict) and set(artifacts) == set(expected_keys), f"{phase}/{name}: artifact inventory mismatch")
    paths: dict[str, Path] = {}
    for key in expected_keys:
        metadata = artifacts[key]
        require(isinstance(metadata, dict), f"{phase}/{name}: invalid {key} artifact metadata")
        relative = text(metadata.get("path"), f"{phase}/{name}.{key}.path")
        path = safe_bundle(relative)
        expected = HERE / phase / f"{name}{suffixes[key]}"
        require(path == expected.resolve(), f"{phase}/{name}: unexpected {key} artifact path")
        expected_bytes = integer(metadata.get("bytes"), f"{phase}/{name}.{key}.bytes")
        expected_sha = digest(metadata.get("sha256"), f"{phase}/{name}.{key}.sha256")
        require(path.stat().st_size == expected_bytes, f"{phase}/{name}: {key} artifact size mismatch")
        require(sha_file(path) == expected_sha, f"{phase}/{name}: {key} artifact hash mismatch")
        paths[key] = path
    return paths


def expected_core_names(protocol: dict[str, Any]) -> list[str]:
    return [
        f"{lane['corpus']}-{lane['provider_label']}-{lane['repeat'].lower()}"
        for lane in protocol["core_order"]
    ]


def expected_core_lane(protocol: dict[str, Any], name: str) -> dict[str, Any]:
    for lane in protocol["core_order"]:
        candidate = f"{lane['corpus']}-{lane['provider_label']}-{lane['repeat'].lower()}"
        if candidate == name:
            return lane
    fail(f"unknown core lane: {name}")


def validate_core_phase(
    phase: str,
    protocol: dict[str, Any],
    freeze: dict[str, Any],
    verifier: Any,
) -> dict[str, dict[str, Any]]:
    folder = HERE / phase
    require(folder.is_dir(), f"missing phase directory: {phase}")
    expected_names = set(expected_core_names(protocol))
    receipts = sorted(folder.glob("*.receipt.json"))
    actual_names = {path.name.removesuffix(".receipt.json") for path in receipts}
    require(actual_names == expected_names, f"{phase}: expected exactly eight core receipts, found {sorted(actual_names)}")
    records: dict[str, dict[str, Any]] = {}
    for name in sorted(expected_names):
        receipt_path = folder / f"{name}.receipt.json"
        receipt = load_json(receipt_path)
        lane = expected_core_lane(protocol, name)
        require(receipt.get("change") == CHANGE and receipt.get("phase") == phase, f"{receipt_path}: identity mismatch")
        require(receipt.get("status") == "pass" and receipt.get("exit_code") == 0, f"{receipt_path}: capture did not pass")
        require(receipt.get("cleanup_verified") is True, f"{receipt_path}: cleanup was not verified")
        require(receipt.get("name") == name and receipt.get("lane") == lane, f"{receipt_path}: lane mismatch")
        require(receipt.get("binary") == freeze["binary"], f"{receipt_path}: binary receipt differs from freeze")
        require(receipt.get("revision") == freeze["revision"], f"{receipt_path}: revision differs from freeze")
        require(receipt.get("protocol_sha256") == freeze["protocol_sha256"], f"{receipt_path}: protocol hash differs")
        require(receipt.get("source_manifest_sha256") == freeze["source_manifest"]["sha256"], f"{receipt_path}: source hash differs")
        require(receipt.get("capture_sha256") == freeze["bound_files"]["capture.py"], f"{receipt_path}: capture driver hash differs")
        require(receipt.get("verifier_sha256") == freeze["bound_files"]["verify-report.py"], f"{receipt_path}: report verifier hash differs")
        require(receipt.get("source_unchanged") is True and receipt.get("source_after") == freeze["source_manifest"], f"{receipt_path}: source custody mismatch")
        command = receipt.get("command")
        require(isinstance(command, list), f"{receipt_path}: missing command")
        report_path = HERE / phase / f"{name}.report.json"
        resource_path = HERE / phase / f"{name}.time.txt"
        expected_command = [
            "taskset",
            "-c",
            protocol["cpu_set"],
            "/usr/bin/time",
            "-v",
            "-o",
            str(resource_path.resolve()),
            freeze["binary"]["path"],
            "provider-lifecycle",
            "--corpus",
            lane["corpus"],
            "--provider",
            lane["provider"],
            "--samples",
            str(protocol["samples"]),
            "--warmup",
            str(protocol["warmups"]),
            "--source-revision",
            freeze["revision"],
            "--output",
            str(report_path.resolve()),
        ]
        if lane["provider"] == "range":
            fail(f"{phase}/{name}: range lane unexpectedly included in core evidence")
        require(command == expected_command, f"{receipt_path}: command differs from protocol")
        tmpdir = receipt.get("environment", {}).get("TMPDIR")
        require(isinstance(tmpdir, str) and Path(tmpdir).resolve().is_relative_to(TMP_ROOT), f"{receipt_path}: TMPDIR is outside owned namespace")
        filesystem = receipt.get("tmpdir_filesystem")
        require(isinstance(filesystem, dict) and "tmpfs" in str(filesystem.get("statfs_type", "")).lower(), f"{receipt_path}: TMPDIR is not recorded as tmpfs")
        artifact_paths = validate_artifacts(
            receipt,
            phase,
            name,
            CORE_ARTIFACTS,
            {"report": ".report.json", "resource": ".time.txt", "stdout": ".stdout.txt", "oracle": ".oracle.txt"},
        )
        try:
            report = verifier.load(artifact_paths["report"])
            checked = checked_report(verifier, report, protocol["samples"], protocol["warmups"])
        except Exception as error:
            fail(f"{phase}/{name}: independent report verifier rejected report: {error}")
        require(report.get("source_revision") == freeze["revision"], f"{phase}/{name}: report revision mismatch")
        require(report.get("binary_sha256") == freeze["binary"]["sha256"] and report.get("binary_bytes") == freeze["binary"]["bytes"], f"{phase}/{name}: report binary mismatch")
        require(report.get("current_exe") == freeze["binary"]["path"], f"{phase}/{name}: report executable mismatch")
        require(checked["corpus"] == lane["corpus"] and checked["provider"] == lane["provider"], f"{phase}/{name}: report provider mismatch")
        oracle_lines = artifact_paths["oracle"].read_text().splitlines()
        require(len(oracle_lines) == 1, f"{phase}/{name}: oracle must contain one JSON line")
        oracle = load_json(artifact_paths["oracle"]) if False else None
        try:
            oracle = json.loads(oracle_lines[0], parse_constant=lambda value: fail(f"{phase}/{name}: non-finite oracle"))
        except (json.JSONDecodeError, CheckError) as error:
            fail(f"{phase}/{name}: invalid oracle: {error}")
        require(oracle == {"status": "pass", **checked}, f"{phase}/{name}: oracle does not exactly match independently checked report")
        identity = report_identity(report)
        records[name] = {
            "name": name,
            "lane": lane,
            "receipt": receipt,
            "report": report,
            "identity": identity,
            "output_identity": output_identity(report),
            "counters": counter_fingerprint(report),
            "counter_sha256": sha_bytes(canonical_bytes(counter_fingerprint(report))),
            "rss_bytes": resource_rss(artifact_paths["resource"]),
            "report_sha256": sha_file(artifact_paths["report"]),
            "oracle_sha256": sha_file(artifact_paths["oracle"]),
            "resource_sha256": sha_file(artifact_paths["resource"]),
            "samples": report["samples"],
            "warmups": report["warmup"],
        }
    require(sum(row["samples"] for row in records.values()) == protocol["core_retained_samples"], f"{phase}: measured sample total mismatch")
    require(sum(row["warmups"] for row in records.values()) == len(records) * protocol["warmups"], f"{phase}: warmup total mismatch")
    return records


def verify_core_relationship(
    protocol: dict[str, Any],
    before: dict[str, dict[str, Any]],
    after: dict[str, dict[str, Any]],
) -> None:
    require(set(before) == set(after) == set(expected_core_names(protocol)), "before/after core lane inventory differs")
    for name in sorted(before):
        left, right = before[name], after[name]
        require(left["output_identity"] == right["output_identity"], f"{name}: before/after output identity differs")
        require(left["counters"] == right["counters"], f"{name}: before/after source/cache/budget counters differ")
    for corpus in ("plain", "media-rich"):
        identities = {
            row["output_identity"]
            for row in list(before.values()) + list(after.values())
            if row["lane"]["corpus"] == corpus
        }
        require(len(identities) == 1, f"{corpus}: provider/repeat output identities are not equal")


def percentile(values: list[int], fraction: float) -> float:
    ordered = sorted(values)
    require(bool(ordered), "cannot summarize empty distribution")
    position = (len(ordered) - 1) * fraction
    left = int(position)
    right = min(left + 1, len(ordered) - 1)
    return ordered[left] + (ordered[right] - ordered[left]) * (position - left)


def distribution(values: list[int]) -> dict[str, Any]:
    return {
        "samples": len(values),
        "min": min(values),
        "p50": percentile(values, 0.50),
        "p95": percentile(values, 0.95),
        "p99": percentile(values, 0.99),
        "max": max(values),
        "mean": statistics.fmean(values),
    }


def relative(before: float, after: float) -> float | None:
    return None if before == 0 else (after - before) * 100.0 / before


def recompute_comparison(protocol: dict[str, Any], before: dict[str, dict[str, Any]], after: dict[str, dict[str, Any]]) -> dict[str, Any]:
    comparisons = []
    flags = []
    for name in sorted(before):
        left, right = before[name], after[name]
        if left["identity"] != right["identity"]:
            flags.append({"lane": name, "metric": "identity", "kind": "oracle", "before": left["identity"], "after": right["identity"]})
        if left["counters"] != right["counters"]:
            flags.append({"lane": name, "metric": "source_cache_budget_counters", "kind": "oracle", "before_after_equal": False})
        for metric in TIMINGS:
            before_values = [row["timings"][metric] for row in left["report"]["samples_raw"]]
            after_values = [row["timings"][metric] for row in right["report"]["samples_raw"]]
            bdist, adist = distribution(before_values), distribution(after_values)
            for statistic in ("p50", "p95", "p99", "mean"):
                change = relative(bdist[statistic], adist[statistic])
                row = {
                    "lane": name,
                    "metric": metric,
                    "statistic": statistic,
                    "before": bdist[statistic],
                    "after": adist[statistic],
                    "relative_percent": change,
                }
                comparisons.append(row)
                if change is not None and abs(change) > protocol["review_percent"]:
                    flags.append({**row, "kind": "timing", "threshold_percent": protocol["review_percent"]})
        bthroughput = [left["report"]["expected_output_bytes"] * 1_000_000_000 / row["timings"]["api_sum_ns"] for row in left["report"]["samples_raw"]]
        athroughput = [right["report"]["expected_output_bytes"] * 1_000_000_000 / row["timings"]["api_sum_ns"] for row in right["report"]["samples_raw"]]
        bd, ad = distribution([int(value) for value in bthroughput]), distribution([int(value) for value in athroughput])
        change = relative(bd["p50"], ad["p50"])
        row = {
            "lane": name,
            "metric": "throughput_bytes_per_second",
            "statistic": "p50",
            "before": bd["p50"],
            "after": ad["p50"],
            "relative_percent": change,
        }
        comparisons.append(row)
        if change is not None and abs(change) > protocol["review_percent"]:
            flags.append({**row, "kind": "throughput", "threshold_percent": protocol["review_percent"]})
        rss_change = relative(left["rss_bytes"], right["rss_bytes"])
        rss_row = {
            "lane": name,
            "metric": "whole_child_rss_bytes",
            "statistic": "endpoint",
            "before": left["rss_bytes"],
            "after": right["rss_bytes"],
            "relative_percent": rss_change,
        }
        comparisons.append(rss_row)
        if rss_change is not None and abs(rss_change) > protocol["review_percent"]:
            flags.append({**rss_row, "kind": "rss", "threshold_percent": protocol["review_percent"]})
    return {
        "change": CHANGE,
        "lanes": sorted(before),
        "before_reports": len(before),
        "after_reports": len(after),
        "samples_per_report": protocol["samples"],
        "comparisons": comparisons,
        "flags": flags,
        "review_percent": protocol["review_percent"],
        "claims": protocol["claims"],
    }


def verify_comparison(protocol: dict[str, Any], before: dict[str, dict[str, Any]], after: dict[str, dict[str, Any]]) -> dict[str, Any]:
    expected = recompute_comparison(protocol, before, after)
    path = HERE / "comparison.json"
    actual = load_json(path)
    require(actual == expected, "comparison.json differs from independent recomputation")
    return {
        "path": str(path.relative_to(HERE)),
        "sha256": sha_file(path),
        "lanes": len(expected["lanes"]),
        "comparisons": len(expected["comparisons"]),
        "flags": len(expected["flags"]),
        "flag_kinds": sorted({row["kind"] for row in expected["flags"]}),
    }


def verify_default_baseline(protocol: dict[str, Any], after_freeze: dict[str, Any]) -> dict[str, Any]:
    """Verify the two retained 37-case/201-row default baseline repeats."""
    summary_path = HERE / "baseline-capture-summary.json"
    summary = load_json(summary_path)
    require(summary.get("schema") == "litchi-0501-default-baseline-summary-v1", "default baseline summary schema mismatch")
    require(summary.get("candidate_base_revision") == after_freeze["revision"], "default baseline candidate revision mismatch")
    require(summary.get("binary") == after_freeze["binary"], "default baseline summary binary mismatch")
    driver = summary.get("driver")
    require(isinstance(driver, dict), "default baseline summary driver is missing")
    driver_path = safe_repo(text(driver.get("path"), "default baseline driver.path"))
    require(sha_file(driver_path) == digest(driver.get("sha256"), "default baseline driver.sha256"), "default baseline driver hash mismatch")
    require(integer(driver.get("bytes"), "default baseline driver.bytes") == driver_path.stat().st_size, "default baseline driver size mismatch")
    lanes = summary.get("lanes")
    require(isinstance(lanes, list) and {row.get("lane") for row in lanes if isinstance(row, dict)} == {"R1-normal", "R2-normal"}, "default baseline must contain R1 and R2")
    require(len(lanes) == 2, "default baseline repeat count differs")
    selected_revision, selected_source = selected_source_files(HERE / "after-source.json")
    require(selected_revision is None or selected_revision == after_freeze["revision"], "default baseline source revision differs")
    summary_source = summary.get("source_manifest")
    require(isinstance(summary_source, dict) and summary_source.get("files") == selected_source, "default baseline summary source differs from after-source")
    summary_source_artifact = summary_source.get("artifact")
    require(isinstance(summary_source_artifact, dict), "default baseline summary source artifact is missing")
    require(digest(summary_source_artifact.get("sha256"), "default baseline summary source sha256") == sha_file(HERE / "after-source.json"), "default baseline summary source artifact hash differs")
    static_tests = summary.get("static_coverage_tests")
    require(isinstance(static_tests, dict) and static_tests.get("exit_code") == 0, "default baseline static coverage tests did not pass")
    static_argv = static_tests.get("argv")
    require(isinstance(static_argv, list) and len(static_argv) == 5 and Path(str(static_argv[0])).name == "python3", "default baseline static test command interpreter differs")
    require(static_argv[1:] == ["-B", "-m", "unittest", "tools.test_crud_coverage_index"], "default baseline static test command differs")
    static_artifact = static_tests.get("artifact")
    require(isinstance(static_artifact, dict), "default baseline static test artifact is missing")
    static_path = safe_repo(text(static_artifact.get("path"), "default baseline static test artifact.path"))
    require(static_path == (HERE / "coverage-tests.log").resolve(), "default baseline static test artifact path differs")
    require(integer(static_artifact.get("bytes"), "default baseline static test artifact.bytes") == static_path.stat().st_size, "default baseline static test artifact size differs")
    require(digest(static_artifact.get("sha256"), "default baseline static test artifact.sha256") == sha_file(static_path), "default baseline static test artifact hash differs")
    for key in ("source_before", "source_after"):
        static_source = static_tests.get(key)
        require(isinstance(static_source, dict) and static_source.get("files") == selected_source, f"default baseline static test {key} differs from after-source")
    temporary_directory = summary.get("temporary_directory")
    require(isinstance(temporary_directory, dict) and temporary_directory.get("absent") is True and not Path(text(temporary_directory.get("path"), "default baseline summary temporary_directory.path")).exists(), "default baseline summary temporary directory was not cleaned")
    lane_summaries = {}
    total_rows = 0
    total_warmups = 0
    corpus_catalog_hashes = set()
    for summary_row in lanes:
        require(isinstance(summary_row, dict), "default baseline summary lane is not an object")
        lane_name = text(summary_row.get("lane"), "default baseline lane")
        require(lane_name in {"R1-normal", "R2-normal"} and lane_name not in lane_summaries, "default baseline lane inventory mismatch")
        lane_dir = HERE / "captures" / lane_name
        receipt_path = lane_dir / "receipt.json"
        receipt = load_json(receipt_path)
        require(summary_row == receipt, f"default baseline summary does not equal {lane_name} receipt")
        require(receipt.get("schema") == "litchi-0501-default-baseline-v1", f"{lane_name}: baseline receipt schema mismatch")
        require(receipt.get("status") == "pass" and receipt.get("exit_code") == 0, f"{lane_name}: baseline receipt failed")
        require(receipt.get("candidate_base_revision") == after_freeze["revision"], f"{lane_name}: candidate revision mismatch")
        require(receipt.get("binary") == after_freeze["binary"] and receipt.get("binary_after") == after_freeze["binary"], f"{lane_name}: binary identity mismatch")
        require(receipt.get("driver_sha256") == driver["sha256"], f"{lane_name}: driver hash differs from summary")
        require(receipt.get("samples") == 15 and receipt.get("warmups") == 3 and receipt.get("expected_rows") == 201, f"{lane_name}: default dimensions differ")
        require(receipt.get("source_unchanged") is True, f"{lane_name}: source changed during default baseline")
        source_manifest = receipt.get("source_manifest")
        require(isinstance(source_manifest, dict), f"{lane_name}: source manifest missing")
        require(source_manifest.get("files") == selected_source and source_manifest.get("files_count") == len(selected_source), f"{lane_name}: source manifest differs from after-source")
        manifest_artifact = source_manifest.get("artifact")
        require(isinstance(manifest_artifact, dict), f"{lane_name}: source manifest artifact missing")
        manifest_path = safe_repo(text(manifest_artifact.get("path"), f"{lane_name}.source_manifest.path"))
        require(manifest_path == (HERE / "after-source.json").resolve(), f"{lane_name}: source manifest path differs")
        require(digest(manifest_artifact.get("sha256"), f"{lane_name}.source_manifest.sha256") == sha_file(manifest_path), f"{lane_name}: source manifest artifact hash differs")
        require(integer(manifest_artifact.get("bytes"), f"{lane_name}.source_manifest.bytes") == manifest_path.stat().st_size, f"{lane_name}: source manifest artifact size differs")
        for key in ("source_before", "source_after"):
            custody = receipt.get(key)
            require(isinstance(custody, dict) and custody.get("files") == selected_source, f"{lane_name}: {key} differs from after-source")
            require(custody.get("manifest") == manifest_artifact, f"{lane_name}: {key} manifest receipt differs")
        cleanup = receipt.get("temporary_directory_cleanup")
        require(isinstance(cleanup, dict) and cleanup.get("absent") is True and not Path(text(cleanup.get("path"), f"{lane_name}.cleanup.path")).exists(), f"{lane_name}: default temporary directory was not cleaned")
        artifacts = receipt.get("artifacts")
        expected_artifact_names = {"report.json", "corpus-catalog.json", "resource.log", "started.json", "stdout.log", "stderr.log", "validation.log"}
        require(isinstance(artifacts, dict) and set(artifacts) == expected_artifact_names, f"{lane_name}: default artifact inventory mismatch")
        resolved_artifacts = {}
        for artifact_name, metadata in artifacts.items():
            require(isinstance(metadata, dict), f"{lane_name}: invalid {artifact_name} metadata")
            path_string = text(metadata.get("path"), f"{lane_name}.{artifact_name}.path")
            path = safe_repo(path_string)
            expected_path = lane_dir / artifact_name
            require(path == expected_path.resolve(), f"{lane_name}: unexpected {artifact_name} path")
            require(integer(metadata.get("bytes"), f"{lane_name}.{artifact_name}.bytes") == path.stat().st_size, f"{lane_name}: {artifact_name} size mismatch")
            require(digest(metadata.get("sha256"), f"{lane_name}.{artifact_name}.sha256") == sha_file(path), f"{lane_name}: {artifact_name} hash mismatch")
            resolved_artifacts[artifact_name] = path
        expected_argv = [
            "/usr/bin/time",
            "-v",
            "-o",
            str((HERE / "captures" / lane_name / "resource.log").relative_to(REPO)),
            "taskset",
            "-c",
            "2",
            after_freeze["binary"]["path"],
            "--workers",
            "1",
            "--samples",
            "15",
            "--warmup",
            "3",
            "--json",
            str((HERE / "captures" / lane_name / "report.json").relative_to(REPO)),
            "--corpus-manifest",
            str((HERE / "captures" / lane_name / "corpus-catalog.json").relative_to(REPO)),
        ]
        require(receipt.get("argv") == expected_argv, f"{lane_name}: baseline command differs")
        validation = receipt.get("validation")
        require(isinstance(validation, dict) and validation.get("exit_code") == 0, f"{lane_name}: validation receipt failed")
        validation_argv = validation.get("argv")
        require(isinstance(validation_argv, list) and len(validation_argv) == 7 and Path(str(validation_argv[0])).name == "python3", f"{lane_name}: validation command interpreter differs")
        require(validation_argv[1:] == ["-B", "tools/validate_crud_coverage_index.py", "--catalog", str((HERE / "captures" / lane_name / "corpus-catalog.json").relative_to(REPO)), "--report", str((HERE / "captures" / lane_name / "report.json").relative_to(REPO))], f"{lane_name}: validation command differs")
        require(validation.get("catalog") == artifacts["corpus-catalog.json"], f"{lane_name}: validation catalog receipt differs")
        require(validation.get("report") == artifacts["report.json"], f"{lane_name}: validation report receipt differs")
        require(resolved_artifacts["validation.log"].read_text() == validation.get("output"), f"{lane_name}: validation log differs from receipt")

        catalog = load_json(resolved_artifacts["corpus-catalog.json"])
        require(catalog.get("manifest_version") == 2 and catalog.get("manifest_kind") == "corpus-catalog", f"{lane_name}: corpus catalog identity differs")
        require(isinstance(catalog.get("corpora"), list) and catalog["corpora"], f"{lane_name}: corpus catalog has no corpora")
        require(isinstance(catalog.get("case_bindings"), list) and len(catalog["case_bindings"]) == 201, f"{lane_name}: corpus catalog binding count differs")
        corpus_catalog_hashes.add(catalog.get("catalog_sha256"))
        report = load_json(resolved_artifacts["report.json"])
        require(report.get("schema_version") == 1 and report.get("tool", {}).get("name") == "litchi-perf-baseline", f"{lane_name}: default report schema/tool differs")
        configuration = report.get("configuration")
        require(isinstance(configuration, dict), f"{lane_name}: default configuration missing")
        require(configuration.get("samples_per_case") == 15 and configuration.get("warmup_iterations_per_case") == 3, f"{lane_name}: default sample dimensions differ")
        require(configuration.get("execution_workers") == [1] and configuration.get("filesystem_process_isolated") is True and configuration.get("filesystem_fresh_child_per_sample") is True, f"{lane_name}: default isolation configuration differs")
        require(isinstance(configuration.get("cases"), list) and len(configuration["cases"]) == 37, f"{lane_name}: default case count differs")
        results = report.get("results")
        require(isinstance(results, list) and len(results) == 201, f"{lane_name}: default result count differs")
        keys = set()
        for index, row in enumerate(results):
            require(isinstance(row, dict) and isinstance(row.get("case"), str) and isinstance(row.get("corpus"), dict), f"{lane_name}: malformed default result {index}")
            corpus = row["corpus"]
            key = (row["case"], corpus.get("package_format"), corpus.get("archive_sha256"))
            require(key not in keys, f"{lane_name}: duplicate default result identity {key}")
            keys.add(key)
            elapsed = row.get("elapsed_ns")
            require(isinstance(elapsed, dict) and isinstance(elapsed.get("samples"), list) and len(elapsed["samples"]) == 15, f"{lane_name}: malformed timing samples {index}")
            require(all(isinstance(value, int) and not isinstance(value, bool) and value >= 0 for value in elapsed["samples"]), f"{lane_name}: invalid timing sample {index}")
        binary_identity = report.get("binary_identity")
        require(isinstance(binary_identity, dict), f"{lane_name}: default binary identity missing")
        require(binary_identity.get("path") == after_freeze["binary"]["path"] and binary_identity.get("binary_sha256") == after_freeze["binary"]["sha256"] and binary_identity.get("binary_bytes") == after_freeze["binary"]["bytes"], f"{lane_name}: default report binary differs")
        reference = report.get("corpus_catalog")
        require(isinstance(reference, dict), f"{lane_name}: default report catalog reference missing")
        for key in ("manifest_version", "catalog_id", "catalog_sha256", "content_set_sha256"):
            require(reference.get(key) == catalog.get(key), f"{lane_name}: report/catalog {key} differs")
        total_rows += len(results) * 15
        total_warmups += len(results) * 3
        lane_summaries[lane_name] = {
            "receipt_sha256": sha_file(receipt_path),
            "report_sha256": sha_file(resolved_artifacts["report.json"]),
            "catalog_sha256": sha_file(resolved_artifacts["corpus-catalog.json"]),
            "rows": len(results),
            "samples": 15,
            "warmups": 3,
        }
    require(len(corpus_catalog_hashes) == 1, "default baseline repeats used different catalog identities")
    require(total_rows == 402 * 15 and total_warmups == 402 * 3, "default baseline total sample count differs")
    return {
        "summary_sha256": sha_file(summary_path),
        "driver_sha256": driver["sha256"],
        "lanes": lane_summaries,
        "rows": total_rows,
        "warmups": total_warmups,
        "catalog_sha256": next(iter(corpus_catalog_hashes)),
    }


def profile_lanes(protocol: dict[str, Any]) -> list[dict[str, str]]:
    result = []
    for corpus in protocol["profile"]["corpora"]:
        for provider in protocol["profile"]["providers"]:
            result.append(
                {
                    "corpus": corpus,
                    "provider": provider,
                    "provider_label": "owned" if provider == "bytes" else "file-warm",
                }
            )
    return result


def walk_dicts(value: Any) -> Iterable[dict[str, Any]]:
    if isinstance(value, dict):
        yield value
        for child in value.values():
            yield from walk_dicts(child)
    elif isinstance(value, list):
        for child in value:
            yield from walk_dicts(child)


def custody_files() -> list[Path]:
    result = [
        HERE / "profile-custody.json",
        HERE / "profile-export-custody.json",
        HERE / "profiles/profile-custody.json",
        HERE / "profiles/profile-export-custody.json",
    ]
    result.extend(HERE.glob("*custody*.json"))
    result.extend((HERE / "profiles").glob("*custody*.json"))
    return list(dict.fromkeys(path for path in result if path.is_file()))


def retained_path(value: str) -> Path:
    path = Path(value)
    if path.is_absolute():
        path = path.resolve()
        require(path.is_relative_to(TMP_ROOT), f"profile export escapes owned tmp root: {path}")
        return path
    return safe_bundle(value)


def recovered_profile_custody(phase: str, name: str, receipt: dict[str, Any]) -> dict[str, Any]:
    expected_sha = receipt.get("raw_perf_sha256")
    expected_bytes = receipt.get("raw_perf_bytes")
    require(isinstance(expected_sha, str) and isinstance(expected_bytes, int), f"{phase}/{name}: missing raw perf identity")
    for manifest_path in custody_files():
        manifest = load_json(manifest_path)
        for node in walk_dicts(manifest):
            strings = " ".join(str(value) for value in node.values() if isinstance(value, str))
            if phase not in strings or name not in strings:
                continue
            sha = None
            size = None
            exported = None
            same_data = None
            raw_removed = None
            resampled = None
            for key, value in node.items():
                lower = str(key).lower()
                if lower in {"raw_perf_sha256", "raw_sha256", "source_sha256", "sha256"} and isinstance(value, str):
                    sha = value
                elif lower in {"raw_perf_bytes", "raw_bytes", "source_bytes", "bytes"} and isinstance(value, int):
                    size = value
                elif lower in {"exported_path", "export_path", "retained_path", "path", "archive"} and isinstance(value, str):
                    exported = value
                elif lower in {"same_data", "same_recording", "raw_data_same"} and isinstance(value, bool):
                    same_data = value
                elif lower in {"raw_perf_was_rerecorded", "rerecorded"} and isinstance(value, bool):
                    same_data = not value
                elif lower in {"raw_removed", "raw_deleted", "original_removed", "source_removed"} and isinstance(value, bool):
                    raw_removed = value
                elif lower in {"resampled", "sampling_repeated"} and isinstance(value, bool):
                    resampled = value
            if sha != expected_sha or size != expected_bytes or exported is None:
                continue
            require(same_data is True, f"{manifest_path}: recovered profile data is not explicitly same-data")
            require(raw_removed is True, f"{manifest_path}: original raw profile removal is not explicit")
            require(resampled is not True, f"{manifest_path}: recovered profile data was resampled")
            exported_path = retained_path(exported)
            require(exported_path.is_file(), f"{manifest_path}: missing retained profile export {exported_path}")
            require(exported_path.stat().st_size == expected_bytes, f"{manifest_path}: recovered profile size mismatch")
            require(sha_file(exported_path) == expected_sha, f"{manifest_path}: recovered profile hash mismatch")
            return {
                "mode": "recovered-export",
                "manifest": str(manifest_path.relative_to(HERE)),
                "path": str(exported_path),
                "sha256": expected_sha,
                "bytes": expected_bytes,
            }
    fail(f"{phase}/{name}: raw perf is absent and no honest same-data custody record exists")


def verify_profile_phase(phase: str, protocol: dict[str, Any], freeze: dict[str, Any], verifier: Any) -> dict[str, Any]:
    folder = HERE / "profiles" / phase
    require(folder.is_dir(), f"missing profile directory: {phase}")
    expected = {f"{lane['corpus']}-{lane['provider_label']}" for lane in profile_lanes(protocol)}
    receipt_paths = sorted(folder.glob("*.receipt.json"))
    actual = {path.name.removesuffix(".receipt.json") for path in receipt_paths}
    require(actual == expected, f"{phase}: profile inventory mismatch")
    summary = load_json(folder / "summary.json")
    require(summary.get("change") == CHANGE and summary.get("phase") == phase, f"{phase}: profile summary identity mismatch")
    require(summary.get("scope") == protocol["profile"]["scope"], f"{phase}: profile scope differs from protocol")
    require(summary.get("sha_hotness_requirement") is True, f"{phase}: profile SHA hotness requirement was relaxed")
    summary_rows = summary.get("profiles")
    require(isinstance(summary_rows, list) and len(summary_rows) == len(expected), f"{phase}: profile summary row count mismatch")
    summary_by_name = {
        f"{row.get('lane', {}).get('corpus')}-{row.get('lane', {}).get('provider_label')}": row
        for row in summary_rows
        if isinstance(row, dict)
    }
    require(set(summary_by_name) == expected, f"{phase}: profile summary lane inventory mismatch")

    records: dict[str, Any] = {}
    for name in sorted(expected):
        receipt = load_json(folder / f"{name}.receipt.json")
        lane = next(lane for lane in profile_lanes(protocol) if f"{lane['corpus']}-{lane['provider_label']}" == name)
        require(receipt.get("change") == CHANGE and receipt.get("phase") == phase and receipt.get("status") == "pass", f"{phase}/{name}: profile receipt failed")
        require(receipt.get("lane") == lane, f"{phase}/{name}: profile lane mismatch")
        require(receipt.get("profile") == protocol["profile"], f"{phase}/{name}: profile protocol differs")
        require(receipt.get("binary") == freeze["binary"], f"{phase}/{name}: profile binary differs")
        require(receipt.get("revision") == freeze["revision"], f"{phase}/{name}: profile revision differs")
        require(receipt.get("protocol_sha256") == freeze["protocol_sha256"], f"{phase}/{name}: profile protocol hash differs")
        require(receipt.get("source_manifest_sha256") == freeze["source_manifest"]["sha256"], f"{phase}/{name}: profile source hash differs")
        artifact_paths = validate_artifacts(receipt, "profiles/" + phase, name, PROFILE_ARTIFACTS, PROFILE_SUFFIXES)
        # validate_artifacts expects a phase directory under HERE.  Profile
        # artifacts use profiles/<phase>, so repeat the path check explicitly.
        for key in PROFILE_ARTIFACTS:
            metadata = receipt["artifacts"][key]
            path = safe_bundle(metadata["path"])
            expected_path = folder / f"{name}{PROFILE_SUFFIXES[key]}"
            require(path == expected_path.resolve(), f"{phase}/{name}: unexpected profile artifact path")
            artifact_paths[key] = path
        try:
            stat_report = verifier.load(artifact_paths["report_json"])
            record_report = verifier.load(artifact_paths["record_report"])
            stat_checked = checked_report(verifier, stat_report, protocol["profile"]["samples"], protocol["profile"]["warmups"])
            record_checked = checked_report(verifier, record_report, protocol["profile"]["samples"], protocol["profile"]["warmups"])
        except Exception as error:
            fail(f"{phase}/{name}: profile report rejected: {error}")
        for label, report, checked in (("stat", stat_report, stat_checked), ("record", record_report, record_checked)):
            require(report.get("source_revision") == freeze["revision"], f"{phase}/{name}/{label}: revision mismatch")
            require(report.get("binary_sha256") == freeze["binary"]["sha256"] and report.get("binary_bytes") == freeze["binary"]["bytes"], f"{phase}/{name}/{label}: binary mismatch")
            require(report.get("current_exe") == freeze["binary"]["path"], f"{phase}/{name}/{label}: executable mismatch")
            require(checked["corpus"] == lane["corpus"] and checked["provider"] == lane["provider"], f"{phase}/{name}/{label}: lane mismatch")
        expected_identities = {
            "stat": {
                "output_sha256": stat_report["expected_output_sha256"],
                "output_bytes": stat_report["expected_output_bytes"],
                "source_archive_sha256": stat_report["source_archive_sha256"],
                "destination_archive_sha256": stat_report["destination_archive_sha256"],
            },
            "record": {
                "output_sha256": record_report["expected_output_sha256"],
                "output_bytes": record_report["expected_output_bytes"],
                "source_archive_sha256": record_report["source_archive_sha256"],
                "destination_archive_sha256": record_report["destination_archive_sha256"],
            },
        }
        require(receipt.get("report_identities") == expected_identities, f"{phase}/{name}: report identity receipt mismatch")
        require(expected_identities["stat"] == expected_identities["record"], f"{phase}/{name}: stat/record output identity differs")
        report_text = artifact_paths["report"].read_text(errors="replace")
        tokens = receipt.get("sha_hotness_tokens_found")
        require(isinstance(tokens, list) and tokens == sorted(set(tokens)) and tokens, f"{phase}/{name}: no SHA hotness token was retained")
        allowed_tokens = set(protocol["profile"]["sha_hotness_tokens"])
        require(set(tokens) <= allowed_tokens, f"{phase}/{name}: unconfigured SHA token claimed")
        require(receipt.get("sha_hotness_detected") is True, f"{phase}/{name}: SHA hotness gate is false")
        for token in tokens:
            require(re.search(re.escape(token), report_text) is not None, f"{phase}/{name}: missing retained token {token}")
        raw_path = Path(text(receipt.get("raw_perf_path"), f"{phase}/{name}.raw_perf_path")).resolve()
        if raw_path.is_file():
            require(raw_path.is_relative_to(TMP_ROOT), f"{phase}/{name}: raw perf escaped owned tmp namespace")
            require(receipt.get("raw_retained_under_owned_tmp") is True, f"{phase}/{name}: raw perf custody flag is false")
            require(integer(receipt.get("raw_perf_bytes"), f"{phase}/{name}.raw_perf_bytes") == raw_path.stat().st_size, f"{phase}/{name}: raw perf size mismatch")
            require(digest(receipt.get("raw_perf_sha256"), f"{phase}/{name}.raw_perf_sha256") == sha_file(raw_path), f"{phase}/{name}: raw perf hash mismatch")
            custody = {"mode": "retained-tmp", "path": str(raw_path), "sha256": receipt["raw_perf_sha256"], "bytes": receipt["raw_perf_bytes"]}
        else:
            custody = recovered_profile_custody(phase, name, receipt)
        require(summary_by_name[name].get("report_identities") == receipt.get("report_identities"), f"{phase}/{name}: summary identity differs")
        records[name] = {
            "lane": lane,
            "receipt_sha256": sha_file(folder / f"{name}.receipt.json"),
            "stat_report_sha256": sha_file(artifact_paths["report_json"]),
            "record_report_sha256": sha_file(artifact_paths["record_report"]),
            "output_identity": expected_identities["stat"],
            "raw_perf": custody,
            "sha_hotness_tokens": tokens,
        }
    return records


def verify_gate_receipts() -> dict[str, Any]:
    commands = load_json(HERE / "gate-commands.json")
    require(isinstance(commands, dict) and commands, "gate-commands.json: empty")
    checks = HERE / "checks"
    required = sorted(commands)
    rows = []
    for name in required:
        receipt_path = checks / f"{name}.json"
        receipt = load_json(receipt_path)
        expected = list(GATE_CPU_PREFIX) + list(commands[name])
        require(receipt.get("command") == expected, f"gate {name}: command differs from gate-commands.json")
        require(receipt.get("exit_code") == 0, f"gate {name}: latest receipt failed")
        log_path = checks / f"{name}.log"
        expected_sha = digest(receipt.get("log_sha256"), f"gate {name}.log_sha256")
        require(expected_sha == sha_file(log_path), f"gate {name}: log hash mismatch")
        rows.append({"name": name, "receipt_sha256": sha_file(receipt_path), "log_sha256": expected_sha, "command": expected})
    attempts = []
    for path in sorted(checks.glob("*.attempt*.json")):
        receipt = load_json(path)
        attempts.append({"path": str(path.relative_to(HERE)), "sha256": sha_file(path), "exit_code": receipt.get("exit_code")})
    return {"required": rows, "attempts": attempts, "commands_sha256": sha_file(HERE / "gate-commands.json")}


def profile_summary(records: dict[str, dict[str, Any]]) -> dict[str, Any]:
    return {
        name: {
            "receipt_sha256": row["receipt_sha256"],
            "stat_report_sha256": row["stat_report_sha256"],
            "record_report_sha256": row["record_report_sha256"],
            "output_identity": row["output_identity"],
            "raw_perf": row["raw_perf"],
            "sha_hotness_tokens": row["sha_hotness_tokens"],
        }
        for name, row in sorted(records.items())
    }


def core_summary(records: dict[str, dict[str, Any]]) -> dict[str, Any]:
    return {
        name: {
            "report_sha256": row["report_sha256"],
            "oracle_sha256": row["oracle_sha256"],
            "resource_sha256": row["resource_sha256"],
            "counter_sha256": row["counter_sha256"],
            "identity": row["identity"],
            "output_identity": row["output_identity"],
            "rss_bytes": row["rss_bytes"],
            "samples": row["samples"],
            "warmups": row["warmups"],
        }
        for name, row in sorted(records.items())
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--before-only", action="store_true", help="verify historical before evidence without requiring after artifacts or gates")
    parser.add_argument("--output", type=Path, default=HERE / "verification-final.json")
    args = parser.parse_args()
    result: dict[str, Any] = {
        "schema": "change-0501-verification-final-v1",
        "change": CHANGE,
        "status": "failed",
    }
    try:
        protocol = load_json(HERE / "protocol.json")
        require(protocol.get("change") == CHANGE, "protocol change mismatch")
        require(protocol.get("samples") == 30 and protocol.get("warmups") == 3, "formal sample dimensions changed")
        require(protocol.get("core_reports") == 8 and protocol.get("core_retained_samples") == 240, "formal core matrix changed")
        before_freeze, before_source_files, driver = verify_freeze(
            "before",
            protocol,
            None,
            strict_driver_custody=not args.before_only,
        )
        protected = verify_protected_work()
        before_build = verify_build("before", before_freeze, HERE / "before-source.json")
        before_verifier_path = locate_old_helper("verify-report.py", before_freeze["bound_files"]["verify-report.py"])
        require(before_verifier_path is not None, "historical before verify-report.py custody is missing")
        before_verifier = module_from(before_verifier_path, "verify_report_final_before")
        before_records = validate_core_phase("before", protocol, before_freeze, before_verifier)
        before_profiles = verify_profile_phase("before", protocol, before_freeze, module_from(HERE / "verify-report.py", "verify_report_final_profile_before"))
        result.update(
            {
                "protocol_sha256": sha_file(HERE / "protocol.json"),
                "claims": protocol["claims"],
                "scope": {
                    "formal": protocol["claims"],
                    "timing": protocol["timing_scope"],
                    "provider": protocol["provider_scope"],
                    "profiles": protocol["profile"]["scope"],
                    "sha_attribution": "whole-child hotness evidence only; no operation-local or blind SHA attribution",
                },
                "before": {
                    "freeze_sha256": sha_file(HERE / "before-freeze.json"),
                    "source_manifest_sha256": before_freeze["source_manifest"]["sha256"],
                    "binary": before_freeze["binary"],
                    "build": before_build,
                    "core": core_summary(before_records),
                    "profiles": profile_summary(before_profiles),
                },
                "protected_work": protected,
                "driver_provenance": driver,
            }
        )
        if not args.before_only:
            after_freeze, after_source_files, after_driver = verify_freeze(
                "after",
                protocol,
                before_freeze,
                strict_driver_custody=True,
            )
            verify_frozen_relationship(before_freeze, after_freeze)
            after_build = verify_build("after", after_freeze, HERE / "after-source.json")
            after_verifier_path = HERE / "verify-report.py"
            require(sha_file(after_verifier_path) == after_freeze["bound_files"]["verify-report.py"], "after report verifier is not the frozen helper")
            after_verifier = module_from(after_verifier_path, "verify_report_final_after")
            after_records = validate_core_phase("after", protocol, after_freeze, after_verifier)
            verify_core_relationship(protocol, before_records, after_records)
            default_baseline = verify_default_baseline(protocol, after_freeze)
            comparison = verify_comparison(protocol, before_records, after_records)
            after_profiles = verify_profile_phase("after", protocol, after_freeze, after_verifier)
            gates = verify_gate_receipts()
            result.update(
                {
                    "after": {
                        "freeze_sha256": sha_file(HERE / "after-freeze.json"),
                        "source_manifest_sha256": after_freeze["source_manifest"]["sha256"],
                        "binary": after_freeze["binary"],
                        "build": after_build,
                        "core": core_summary(after_records),
                        "profiles": profile_summary(after_profiles),
                    },
                    "default_baseline": default_baseline,
                    "comparison": comparison,
                    "gates": gates,
                    "historical_source": {
                        "before_files": len(before_source_files),
                        "after_files": len(after_source_files),
                        "before_archives": "source-archives/before.json",
                    },
                }
            )
        else:
            result["mode"] = "before-only"
        result["status"] = "pass"
    except (CheckError, OSError, KeyError, TypeError, ValueError) as error:
        result["error"] = str(error)
        result["status"] = "failed"
    output = args.output.resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_bytes(canonical_bytes(result))
    print(json.dumps({"status": result["status"], "output": str(output), **({"error": result["error"]} if "error" in result else {})}, sort_keys=True))
    return 0 if result["status"] == "pass" else 1


if __name__ == "__main__":
    sys.exit(main())
