#!/usr/bin/env python3
"""Validate the draft 0437 ODP evidence lifecycle.

This driver is intentionally a terminal read/derive gate.  It does not start
the ODP workload, rebuild a binary, rewrite the sealed inventory, or infer a
matrix from whatever files happen to be present.  The frozen protocol must
name the receipt sets and the three-role matrix explicitly before this gate
can pass.  That keeps the preparatory six-pilot/six-profile plan distinct
from the 36-report/1080-sample formal matrix.
"""

from __future__ import annotations

import argparse
import importlib.util
import json
from pathlib import Path
import sys
from typing import Any


ROOT = Path(__file__).resolve().parent
CHANGE = 437
MARKER_NAME = "workload-verify.json"
FORMAL_ROLES = ("before-buffered", "after-buffered", "after-streaming")
FORMAL_SHAPES = {"tiny": 64, "medium": 4_096, "large": 8_192}


def module_stem(name: str) -> str:
    if not isinstance(name, str) or not name:
        raise ValueError("lifecycle module name must be a non-empty string")
    candidate = Path(name)
    if (
        candidate.is_absolute()
        or len(candidate.parts) != 1
        or candidate.parts[0] in {"", ".", ".."}
        or ".." in candidate.parts
    ):
        raise ValueError(f"unsafe lifecycle module path: {name!r}")
    stem = candidate.name.removesuffix(".py")
    if not stem or stem in {".", ".."} or "/" in stem or "\\" in stem:
        raise ValueError(f"unsafe lifecycle module path: {name!r}")
    return stem


def module(name: str):
    stem = module_stem(name)
    path = (ROOT / (stem + ".py")).resolve()
    if not path.is_relative_to(ROOT.resolve()):
        raise ValueError(f"lifecycle module escapes the bundle: {name!r}")
    spec = importlib.util.spec_from_file_location(
        "change0437_lifecycle_" + stem, path
    )
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load {stem}.py")
    result = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(result)
    return result


def protocol() -> dict[str, Any]:
    path = ROOT / "protocol.json"
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise ValueError(f"invalid 0437 protocol: {error}") from error
    if not isinstance(value, dict) or value.get("change") != CHANGE:
        raise ValueError("protocol is not the frozen 0437 protocol")
    return value


def required_paths(
    contract: dict[str, Any], key: str, *, expected_count: int, label: str
) -> tuple[str, ...]:
    values = contract.get(key)
    if (
        not isinstance(values, list)
        or any(not isinstance(value, str) or not value for value in values)
    ):
        raise ValueError(f"protocol.{label}.{key} must be a nonempty string list")
    if len(values) != expected_count:
        raise ValueError(
            f"protocol.{label}.{key} must contain {expected_count} paths, found {len(values)}"
        )
    if len(set(values)) != len(values):
        raise ValueError(f"protocol.{label}.{key} contains duplicate paths")
    for value in values:
        path = Path(value)
        if path.is_absolute() or ".." in path.parts:
            raise ValueError(f"unsafe bundle-relative path in protocol: {value}")
    return tuple(values)


def matrix_contract(value: dict[str, Any]) -> dict[str, Any]:
    if value.get("shapes") != FORMAL_SHAPES:
        raise ValueError("protocol.shapes must bind tiny=64, medium=4096, large=8192")
    matrix = value.get("matrix")
    if not isinstance(matrix, dict):
        raise ValueError("protocol.matrix must be an object")
    roles = matrix.get("roles")
    if not isinstance(roles, list) or tuple(roles) != FORMAL_ROLES:
        raise ValueError(
            "protocol.matrix.roles must be the ordered three-role tuple "
            "before-buffered,after-buffered,after-streaming"
        )
    expected = {"formal_reports": 36, "retained_samples": 1080}
    for key, count in expected.items():
        if matrix.get(key) != count:
            raise ValueError(f"protocol.matrix.{key} must be {count}")
    return matrix


def bundle_json_name(v: Any, value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        v.fail(label, "must be a non-empty bundle-relative JSON path")
    path = Path(value)
    if (
        path.is_absolute()
        or ".." in path.parts
        or path.name in {"", ".", ".."}
        or path.suffix != ".json"
    ):
        v.fail(label, "must be a safe bundle-relative JSON path")
    return value


def is_profile_marker(path: Path) -> bool:
    try:
        relative = path.relative_to(ROOT).parts
    except ValueError:
        return False
    if len(relative) < 3 or relative[0] != "profiles" or relative[-1] != MARKER_NAME:
        return False
    if path.read_bytes() != b"VALID\n":
        raise ValueError(f"{path}: profile marker is not exactly VALID")
    return True


def validate_planned(v: Any) -> None:
    planned = v.load(ROOT / "planned-checks.json")
    if not isinstance(planned, dict):
        v.fail("planned-checks.json", "expected an object")
    for key in ("required_pass", "required_review"):
        names = planned.get(key)
        if not isinstance(names, list) or any(not isinstance(name, str) for name in names):
            v.fail(f"planned-checks.json.{key}", "expected a list of bundle-relative paths")
        expected_status = "pass" if key == "required_pass" else "failed"
        for name in names:
            path = v.bundle_path(name, f"planned-checks.{key}")
            receipt = v.load(path, name)
            if receipt.get("status") != expected_status:
                v.fail(name, f"required check has wrong terminal status (expected {expected_status})")


def validate_receipt_artifacts(v: Any, path: Path, row: dict[str, Any]) -> None:
    artifacts = row.get("artifacts")
    if artifacts is None:
        return
    if not isinstance(artifacts, dict):
        v.fail(str(path), "artifact inventory must be an object")
    for name, record in artifacts.items():
        if not isinstance(name, str) or not isinstance(record, dict):
            v.fail(str(path), "artifact inventory entry is malformed")
        value = record if "path" in record else dict(record, path=name)
        # The 0437 outer verifier exposes artifact_data(), not the older
        # artifact_path() helper.  Decode and hash every retained artifact so
        # the lifecycle gate has the same raw/gzip custody semantics.
        v.artifact_data(path, value, f"{path}.artifacts.{name}")


def validate_terminal_receipt(v: Any, path: Path, row: dict[str, Any]) -> None:
    label = str(path)
    if row.get("status") not in {"pass", "failed"}:
        v.fail(label, "receipt is not terminal")
    if "source_before" not in row or "source_after" not in row:
        v.fail(label, "receipt has incomplete source custody")
    v.check_source_manifest(row["source_before"], label + ".source_before")
    v.check_source_manifest(row["source_after"], label + ".source_after")
    if row["source_before"] != row["source_after"] or (
        "source_unchanged" in row and row.get("source_unchanged") is not True
    ):
        v.fail(label, "receipt does not prove unchanged source custody")
    if row.get("log") is not None:
        log = row["log"]
        # The current outer verifier exposes artifact_data(), which also
        # accepts the raw/gzip sidecar convention used by sealed bundles.
        # Keep the older check receipt shape (a record with path/bytes/sha)
        # intact while avoiding a live-path-only check here.
        if not isinstance(log, dict):
            v.fail(label + ".log", "log binding must be an artifact record")
        v.artifact_data(path, log, label + ".log")
    validate_receipt_artifacts(v, path, row)


def validate_source_receipts(v: Any) -> int:
    expected = v.load(ROOT / "expected-checks.json")
    if not isinstance(expected, dict):
        v.fail("expected-checks.json", "expected-checks must be an object")
    actual: dict[str, Any] = {}
    for path in sorted(ROOT.rglob("*.json")):
        if path.name in {"compression.json", "expected-checks.json"}:
            continue
        if is_profile_marker(path):
            continue
        row = v.load(path, str(path))
        if not isinstance(row, dict) or "source_before" not in row:
            continue
        name = path.relative_to(ROOT).with_suffix("").as_posix()
        actual[name] = row.get("status")
        validate_terminal_receipt(v, path, row)
        if path.parent == ROOT / "checks":
            if row.get("driver_sha256") != v.sha(ROOT / "check.py"):
                v.fail(name, "command custody driver differs")
            if row.get("status") == "pass" and row.get("exit_code") != 0:
                v.fail(name, "passing command has nonzero exit")
            if row.get("status") == "pass" and row.get("passed_tests", 1) <= 0:
                v.fail(name, "passing test command ran no tests")
    if actual != expected:
        v.fail("expected-checks.json", "terminal receipt inventory differs")
    return len(actual)


def validate_cleanup_gates(v: Any, stage: str, protocol_value: dict[str, Any]) -> None:
    if stage == "precleanup":
        return

    def bundle_name(value: Any, label: str) -> str:
        if not isinstance(value, str) or not value:
            v.fail(label, "must be a non-empty bundle-relative path")
        path = Path(value)
        if path.is_absolute() or ".." in path.parts:
            v.fail(label, "must be a safe bundle-relative path")
        return value

    precleanup_name = bundle_name(
        protocol_value.get("precleanup_receipt", "checks/precleanup-portable.json"),
        "protocol.precleanup_receipt",
    )
    cleanup_name = bundle_name(
        protocol_value.get("cleanup_receipt", "checks/task-cleanup.json"),
        "protocol.cleanup_receipt",
    )
    inventory_name = bundle_name(
        protocol_value.get("cleanup_inventory", "checks/cleanup-inventory.json"),
        "protocol.cleanup_inventory",
    )
    for name in (precleanup_name, cleanup_name):
        path = v.bundle_path(name, "cleanup lifecycle receipt")
        row = v.load(path, name)
        if row.get("status") != "pass" or row.get("exit_code") != 0:
            v.fail(name, "required cleanup lifecycle check is not passing")

    # The copied final bundle cannot inspect the removed scratch directories,
    # preserved target directories, or the source checkout's GOAL.md.  Bind
    # only the retained cleanup inventory to the frozen protocol values.
    cleanup_paths = protocol_value.get("cleanup_paths")
    cleanup_expected_paths = protocol_value.get("cleanup_expected_paths")
    preserved_paths = protocol_value.get("preserved_paths")
    goal_path = protocol_value.get("goal_path")
    goal_sha256 = protocol_value.get("goal_sha256")
    if (
        not isinstance(cleanup_paths, list)
        or not cleanup_paths
        or any(not isinstance(path, str) or not path for path in cleanup_paths)
        or cleanup_expected_paths != cleanup_paths
        or not isinstance(preserved_paths, list)
        or not preserved_paths
        or any(not isinstance(path, str) or not path for path in preserved_paths)
        or not isinstance(goal_path, str)
        or not goal_path
        or not isinstance(goal_sha256, str)
        or len(goal_sha256) != 64
    ):
        v.fail("protocol.cleanup", "cleanup inventory bindings are malformed")

    inventory_path = v.bundle_path(inventory_name, "cleanup inventory")
    inventory = v.load(inventory_path, inventory_name)
    if not isinstance(inventory, dict):
        v.fail(inventory_name, "cleanup inventory must be an object")
    if inventory.get("change") != CHANGE or inventory.get("status") != "pass":
        v.fail(inventory_name, "cleanup inventory is not a passing 0437 record")
    if inventory.get("protocol") != "protocol.json":
        v.fail(inventory_name, "cleanup inventory does not bind protocol.json")
    if inventory.get("precleanup_receipt") != precleanup_name:
        v.fail(inventory_name, "cleanup inventory precleanup receipt binding differs")
    for key, expected in (
        ("cleanup_paths", cleanup_paths),
        ("cleanup_expected_paths", cleanup_expected_paths),
    ):
        if inventory.get(key) != expected:
            v.fail(inventory_name, f"cleanup inventory {key} binding differs")

    removed = inventory.get("removed")
    if not isinstance(removed, list) or len(removed) != len(cleanup_paths):
        v.fail(inventory_name, "cleanup inventory removed list differs")
    for expected_path, record in zip(cleanup_paths, removed):
        if not isinstance(record, dict) or record.get("path") != expected_path:
            v.fail(inventory_name, "cleanup inventory removed path differs")
        for key in ("regular_files", "regular_file_bytes"):
            count = record.get(key)
            if isinstance(count, bool) or not isinstance(count, int) or count < 0:
                v.fail(inventory_name, f"cleanup inventory {key} is invalid")

    identities = inventory.get("preserved_target_directory_identity")
    if not isinstance(identities, dict) or set(identities) != set(preserved_paths):
        v.fail(inventory_name, "preserved target identity bindings differ")
    for path in preserved_paths:
        identity = identities.get(path)
        if (
            not isinstance(identity, list)
            or len(identity) != 2
            or any(isinstance(value, bool) or not isinstance(value, int) or value < 0 for value in identity)
        ):
            v.fail(inventory_name, f"preserved target identity is invalid for {path}")
    if inventory.get("user_goal_path") != goal_path:
        v.fail(inventory_name, "cleanup inventory GOAL path binding differs")
    if inventory.get("user_goal_sha256") != goal_sha256:
        v.fail(inventory_name, "cleanup inventory GOAL digest binding differs")


def validate_hash_binding(v: Any, path: Path, row: dict[str, Any], path_key: str, hash_key: str) -> None:
    value = row.get(path_key)
    if not isinstance(value, str) or not value:
        v.fail(str(path), f"{path_key} binding is missing")
    bound = v.bundle_path(value, f"{path}.{path_key}")
    expected = row.get(hash_key)
    if expected != v.sha(bound):
        v.fail(str(path), f"{path_key}/{hash_key} binding is stale")


def validate_binary_identity(v: Any, path: Path, row: dict[str, Any]) -> None:
    binary = row.get("binary")
    if not isinstance(binary, dict):
        v.fail(str(path), "receipt has no binary identity")
    binary_path = binary.get("path")
    size = binary.get("bytes")
    digest = binary.get("sha256")
    if not isinstance(binary_path, str) or not binary_path or not Path(binary_path).is_absolute():
        v.fail(str(path), "binary identity path must be absolute")
    if isinstance(size, bool) or not isinstance(size, int) or size <= 0:
        v.fail(str(path), "binary identity byte count is invalid")
    if not isinstance(digest, str) or len(digest) != 64:
        v.fail(str(path), "binary identity digest is invalid")


def validate_receipt_path_binding(
    v: Any, name: str, row: dict[str, Any], kind: str
) -> None:
    """Bind receipt fields to their frozen path grammar.

    Pilot receipts predate an explicit ``attempt`` field, so ``initial`` is
    authoritative from the required path.  Profile receipts already carry
    ``attempt`` and ``preparatory`` fields; both are checked against their
    path without changing any historical receipt schema.
    """

    parts = Path(name).parts
    if not isinstance(row, dict):
        v.fail(name, "receipt must be an object")
    if kind == "pilot":
        if len(parts) != 4 or parts[0] != "pilots" or parts[2] != "initial":
            v.fail(name, "pilot receipt path does not match the frozen grammar")
        role = parts[1]
        expected = {
            f"{mode}-{shape}-receipt.json": (mode, shape)
            for mode in ("normal", "allocator")
            for shape in ("tiny", "medium", "large")
        }
        mode_shape = expected.get(parts[3])
        if role not in FORMAL_ROLES or mode_shape is None:
            v.fail(name, "pilot receipt path has an unknown role, mode, or shape")
        mode, shape = mode_shape
        if row.get("role") != role or row.get("mode") != mode or row.get("shape") != shape:
            v.fail(name, "pilot receipt fields differ from their required path")
        return

    if kind == "preparatory-profile":
        expected_role = "before-buffered"
        expected_attempt = "initial"
        expected_preparatory = True
        if (
            len(parts) != 5
            or parts[0] != "profiles"
            or parts[1] != "preparatory-before-buffered"
            or parts[2] != expected_attempt
            or parts[4] != "receipt.json"
            or parts[3] not in {"stat", "record"}
        ):
            v.fail(name, "preparatory profile path does not match the frozen grammar")
        expected_kind = parts[3]
    elif kind == "formal-profile":
        expected_attempt = "formal"
        expected_preparatory = False
        if (
            len(parts) != 4
            or parts[0] != "profiles"
            or parts[1] not in FORMAL_ROLES
            or parts[3] != "receipt.json"
            or parts[2] not in {"stat", "record"}
        ):
            v.fail(name, "formal profile path does not match the frozen grammar")
        expected_role = parts[1]
        expected_kind = parts[2]
    else:
        raise ValueError(f"unknown receipt path contract: {kind}")

    if (
        row.get("role") != expected_role
        or row.get("kind") != expected_kind
        or row.get("shape") != "large"
        or row.get("attempt") != expected_attempt
        or row.get("preparatory") is not expected_preparatory
    ):
        v.fail(name, "profile receipt fields differ from their required path")


def validate_driver_and_oracle_bindings(
    v: Any, path: Path, row: dict[str, Any], protocol_value: dict[str, Any], kind: str
) -> None:
    """Check pilot/profile custody without requiring external binaries live.

    Preparatory pilots/profiles intentionally bind the retained historical
    ``verify-report.py`` and ``protocol-draft.json``.  Formal profiles bind
    the candidate oracle selected by protocol.oracle.  Both paths and their
    hashes remain explicit in the receipt so a terminal status alone cannot
    authorize a report.
    """

    candidate_pilot = (
        kind == "pilot"
        and isinstance(row.get("oracle_path"), str)
        and row.get("oracle_path") == "verify-report-candidate.py"
    )
    driver_name = (
        "pilot-candidate.py" if candidate_pilot else "pilot.py"
    ) if kind == "pilot" else "profile.py"
    driver_hash = row.get("driver_sha256")
    driver_path = ROOT / driver_name
    if not isinstance(driver_hash, str) or not driver_path.is_file() or driver_hash != v.sha(driver_path):
        v.fail(str(path), f"{kind} driver custody is missing or stale")

    protocol_path_name = row.get("protocol_path")
    if not isinstance(protocol_path_name, str) or not protocol_path_name:
        protocol_path_name = "oracle-protocol.json" if candidate_pilot else (
            "protocol-draft.json" if kind == "pilot" or row.get("preparatory") is True else "protocol.json"
        )
    if candidate_pilot and protocol_path_name != "oracle-protocol.json":
        v.fail(str(path), "candidate pilot must bind oracle-protocol.json")
    if kind == "pilot" and not candidate_pilot and protocol_path_name != "protocol-draft.json":
        v.fail(str(path), "baseline pilot must bind protocol-draft.json")
    protocol_hash_key = "protocol_sha256"
    protocol_file = v.bundle_path(protocol_path_name, f"{path}.protocol_path")
    if row.get(protocol_hash_key) != v.sha(protocol_file):
        v.fail(str(path), "protocol custody is stale")

    if kind == "pilot":
        oracle_name = row.get("oracle_path") or protocol_value.get("preparatory_oracle_path", "verify-report.py")
        oracle_hash_key = "oracle_sha256"
    elif row.get("preparatory") is True:
        oracle_name = protocol_value.get("preparatory_oracle_path", "verify-report.py")
        oracle_hash_key = "oracle_verifier_sha256"
    else:
        oracle = protocol_value.get("oracle")
        if not isinstance(oracle, dict):
            v.fail("protocol.oracle", "formal profile oracle binding is missing")
        oracle_name = oracle.get("verifier_path")
        oracle_hash_key = "oracle_verifier_sha256"
    if not isinstance(oracle_name, str) or not oracle_name:
        v.fail(str(path), "oracle path binding is missing")
    if kind == "pilot" and candidate_pilot and oracle_name != "verify-report-candidate.py":
        v.fail(str(path), "candidate pilot must bind verify-report-candidate.py")
    if kind == "pilot" and not candidate_pilot and oracle_name != "verify-report.py":
        v.fail(str(path), "baseline pilot must bind verify-report.py")
    if kind == "profile" and row.get("oracle_verifier") != oracle_name:
        v.fail(str(path), "profile oracle path differs from the expected protocol binding")
    oracle_path = v.bundle_path(oracle_name, f"{path}.oracle_path")
    if row.get(oracle_hash_key) != v.sha(oracle_path):
        v.fail(str(path), "oracle custody is stale")

    validate_binary_identity(v, path, row)
    if kind == "profile":
        capture_helper_hash = row.get("capture_helper_sha256")
        if capture_helper_hash != v.sha(ROOT / "capture.py"):
            v.fail(str(path), "profile capture helper custody is stale")
    if kind == "pilot":
        validate_hash_binding(v, path, row, "build_receipt", "build_receipt_sha256")
    elif kind == "profile":
        validate_hash_binding(v, path, row, "build_receipt", "build_receipt_sha256")
        validate_hash_binding(v, path, row, "binary_copies_receipt", "binary_copies_receipt_sha256")


    build_file = v.bundle_path(row["build_receipt"], f"{path}.build_receipt")
    build = v.load(build_file)
    if build.get("status") != "pass" or build.get("exit_code") != 0 or build.get("source_before") != build.get("source_after") or build.get("source_unchanged") is not True:
        v.fail(str(path), "retained build receipt is not passing with stable source")
    bound_source = row.get("source_manifest", row.get("source_before"))
    if bound_source != build.get("source_before"):
        v.fail(str(path), "receipt source differs from its executable build")
    if "revision" in row and row["revision"] != build.get("revision"):
        v.fail(str(path), "receipt revision differs from its executable build")
    build_role = "before" if row.get("role") == "before-buffered" else "after"
    copies = v.load(ROOT / build_role / "binary-copies.json")
    mode = row.get("mode", "normal")
    if copies.get(mode) != row.get("binary"):
        v.fail(str(path), "receipt binary differs from retained build copies")


def revalidate_preparatory_report(v: Any, path: Path, row: dict[str, Any], kind: str) -> None:
    # Re-run the report oracle against retained bytes; a VALID log alone is
    # not the portable semantic proof. The baseline keeps its original profile.
    is_baseline = row.get("role") == "before-buffered" and (kind == "pilot" or row.get("preparatory") is True)
    filename = "verify-report.py" if is_baseline else "verify-report-candidate.py"
    spec = importlib.util.spec_from_file_location("odp_lifecycle_report_oracle", ROOT / filename)
    oracle = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(oracle)
    if is_baseline:
        oracle.protocol_path = lambda: ROOT / "protocol-draft.json"
    report_record = next((record if "path" in record else dict(record, path=name)
        for name, record in row["artifacts"].items()
        if ((record.get("path", name).endswith(".json"))
            and "catalog" not in record.get("path", name)
            and not record.get("path", name).endswith("receipt.json"))), None)
    if report_record is None:
        v.fail(str(path), "retained report is missing")
    report_path = v.bundle_path(report_record["path"], f"{path}.report")
    oracle.validate_report(report_path, row.get("mode", "normal"), row["shape"], row["role"],
        samples=3 if kind == "pilot" else 30, warmups=1 if kind == "pilot" else 3)


def validate_report_and_oracle_artifacts(v: Any, path: Path, row: dict[str, Any]) -> None:
    """Require retained report identity and an oracle log that says VALID.

    Pilot/profile receipts are preparatory inputs to the formal bundle.  Their
    terminal status and driver hashes are necessary but insufficient: the
    retained report must bind the receipt's binary descriptor, and the copied
    oracle must have actually accepted that report.  The outer formal verifier
    performs the same checks for the 36 matrix rows; this keeps the six pilot
    and two preparatory profile inputs under the same custody rule.
    """

    artifacts = row.get("artifacts")
    if not isinstance(artifacts, dict):
        v.fail(str(path), "report/oracle custody requires an artifact inventory")
    report_record: dict[str, Any] | None = None
    oracle_record: dict[str, Any] | None = None
    report_label = ""
    oracle_label = ""
    for name, raw_record in artifacts.items():
        if not isinstance(name, str) or not isinstance(raw_record, dict):
            continue
        record = raw_record if "path" in raw_record else dict(raw_record, path=name)
        value = record.get("path")
        if not isinstance(value, str):
            continue
        lowered = value.lower()
        if "oracle" in lowered and (lowered.endswith(".log") or lowered.endswith(".log.gz")):
            oracle_record = record
            oracle_label = f"{path}.artifacts.{name}"
        if lowered.endswith(".json") or lowered.endswith(".json.gz"):
            if "catalog" not in lowered and not lowered.endswith("receipt.json"):
                report_record = record
                report_label = f"{path}.artifacts.{name}"
    if report_record is None:
        v.fail(str(path), "receipt has no retained report artifact")
    if oracle_record is None:
        v.fail(str(path), "receipt has no retained oracle log artifact")
    _, oracle_bytes = v.artifact_data(path, oracle_record, oracle_label)
    if b"stdout=VALID" not in oracle_bytes and oracle_bytes.strip() != b"VALID":
        v.fail(oracle_label, "oracle log does not retain stdout=VALID")
    _, report_bytes = v.artifact_data(path, report_record, report_label)
    try:
        report = json.loads(report_bytes.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as error:
        v.fail(report_label, f"report artifact is not valid JSON: {error}")
    reported = report.get("binary_identity") if isinstance(report, dict) else None
    binary = row.get("binary")
    if not isinstance(reported, dict) or not isinstance(binary, dict):
        v.fail(report_label, "report/binary identity is missing")
    for report_key, receipt_key in (
        ("path", "path"),
        ("binary_sha256", "sha256"),
        ("binary_bytes", "bytes"),
    ):
        if reported.get(report_key) != binary.get(receipt_key):
            v.fail(report_label, f"report {report_key} differs from receipt binary identity")


def validate_named_receipts(
    v: Any,
    names: tuple[str, ...],
    *,
    label: str,
    protocol_value: dict[str, Any],
    kind: str,
) -> int:
    for name in names:
        path = v.bundle_path(name, label)
        row = v.load(path, name)
        if not isinstance(row, dict):
            v.fail(name, "receipt must be an object")
        validate_terminal_receipt(v, path, row)
        if row.get("status") != "pass":
            v.fail(name, f"{kind} receipt is not passing")
        if row.get("exit_code") != 0:
            v.fail(name, f"passing {kind} receipt has nonzero exit")
        if row.get("oracle_exit_code") != 0:
            v.fail(name, f"passing {kind} receipt has nonzero oracle exit")
        validate_driver_and_oracle_bindings(v, path, row, protocol_value, kind)
        validate_report_and_oracle_artifacts(v, path, row)
        revalidate_preparatory_report(v, path, row, kind)
    return len(names)


def validate_pilots_and_profiles(v: Any, value: dict[str, Any]) -> tuple[int, int, int]:
    pilots = value.get("pilots")
    if not isinstance(pilots, dict):
        v.fail("protocol.pilots", "pilot contract is missing")
    pilot_names = required_paths(pilots, "required_receipts", expected_count=18, label="pilots")
    pilot_groups = pilots.get("groups")
    if not isinstance(pilot_groups, dict):
        v.fail("protocol.pilots.groups", "pilot group counts are missing")
    if pilot_groups.get("preparatory_before_buffered") != 6:
        v.fail("protocol.pilots.groups", "preparatory before-buffered pilot count must be 6")
    if pilot_groups.get("candidate_two_roles") != 12:
        v.fail("protocol.pilots.groups", "candidate two-role pilot count must be 12")
    preparatory = value.get("preparatory_profiles")
    if not isinstance(preparatory, dict):
        v.fail("protocol.preparatory_profiles", "preparatory profile contract is missing")
    preparatory_names = required_paths(
        preparatory, "required_receipts", expected_count=2, label="preparatory_profiles"
    )
    formal = value.get("formal_profiles")
    if not isinstance(formal, dict):
        v.fail("protocol.formal_profiles", "formal profile contract is missing")
    formal_names = required_paths(
        formal, "required_receipts", expected_count=6, label="formal_profiles"
    )
    for names, label in (
        (pilot_names, "pilots"),
        (preparatory_names, "preparatory_profiles"),
        (formal_names, "formal_profiles"),
    ):
        for name in names:
            if not name.endswith("receipt.json"):
                v.fail(name, f"{label} path must name a receipt.json")
    pilot_role_counts = {role: 0 for role in FORMAL_ROLES}
    for name in pilot_names:
        pilot_path = v.bundle_path(name, "pilots")
        pilot_row = v.load(pilot_path, name)
        validate_receipt_path_binding(v, name, pilot_row, "pilot")
        role = pilot_row.get("role") if isinstance(pilot_row, dict) else None
        if role not in pilot_role_counts:
            v.fail(name, "pilot role is outside the frozen three-role matrix")
        candidate = pilot_row.get("oracle_path") == "verify-report-candidate.py"
        if role == "before-buffered" and candidate:
            v.fail(name, "before-buffered pilot must retain the historical baseline oracle")
        if role in {"after-buffered", "after-streaming"} and not candidate:
            v.fail(name, "candidate role pilot must bind the candidate oracle")
        pilot_role_counts[role] += 1
    if pilot_role_counts != {role: 6 for role in FORMAL_ROLES}:
        v.fail("protocol.pilots", f"pilot role distribution differs: {pilot_role_counts}")
    for name in preparatory_names:
        profile_row = v.load(v.bundle_path(name, "preparatory_profiles"), name)
        validate_receipt_path_binding(v, name, profile_row, "preparatory-profile")
    for name in formal_names:
        profile_row = v.load(v.bundle_path(name, "formal_profiles"), name)
        validate_receipt_path_binding(v, name, profile_row, "formal-profile")
    validate_named_receipts(
        v,
        pilot_names,
        label="pilots",
        protocol_value=value,
        kind="pilot",
    )
    validate_named_receipts(
        v,
        preparatory_names,
        label="preparatory_profiles",
        protocol_value=value,
        kind="profile",
    )
    validate_named_receipts(
        v,
        formal_names,
        label="formal_profiles",
        protocol_value=value,
        kind="profile",
    )
    return len(pilot_names), len(preparatory_names), len(formal_names)


def validate_summary_shape(v: Any, value: Any, matrix: dict[str, Any]) -> None:
    if not isinstance(value, dict):
        v.fail("summary.json", "summary must be an object")
    actual = value.get("matrix")
    if not isinstance(actual, dict):
        v.fail("summary.json.matrix", "summary matrix must be an object")
    for key in ("formal_reports", "retained_samples"):
        if actual.get(key) != matrix[key]:
            v.fail("summary.json.matrix", f"{key} differs from frozen protocol")


def validate(stage: str) -> dict[str, Any]:
    value = protocol()
    matrix = matrix_contract(value)
    v = module("verify")
    validate_planned(v)
    validate_cleanup_gates(v, stage, value)
    terminal_receipts = validate_source_receipts(v)
    pilots, preparatory_profiles, formal_profiles = validate_pilots_and_profiles(v, value)

    summary_module_name = value.get("summary_module", "summary")
    summary_path_name = value.get("summary_path", "summary.json")
    summary = module(module_stem(summary_module_name))
    summary_path_name = bundle_json_name(v, summary_path_name, "protocol.summary_path")
    retained_value = v.load(v.bundle_path(summary_path_name, "summary_path"), summary_path_name)
    derived_value = summary.derive()
    validate_summary_shape(v, retained_value, matrix)
    validate_summary_shape(v, derived_value, matrix)
    if summary.canonical(retained_value) != summary.canonical(derived_value):
        v.fail(summary_path_name, "retained summary differs from independent derivation")

    preparatory_module_name = value.get("preparatory_summary_module", "before-hypothesis")
    preparatory_path_name = value.get("preparatory_summary_path", "before-hypothesis.json")
    preparatory = module(module_stem(preparatory_module_name))
    preparatory_path_name = bundle_json_name(
        v, preparatory_path_name, "protocol.preparatory_summary_path"
    )
    retained_preparatory = v.load(
        v.bundle_path(preparatory_path_name, "preparatory_summary_path"),
        preparatory_path_name,
    )
    if retained_preparatory != preparatory.derive():
        v.fail(preparatory_path_name, "retained preparatory summary differs from independent derivation")

    decision_spec = value.get("retention_decision")
    if not isinstance(decision_spec, dict):
        v.fail("protocol.retention_decision", "retention decision binding is missing")
    decision_name = module_stem(decision_spec.get("driver"))
    decision_driver = v.bundle_path(decision_name + ".py", "retention decision driver")
    if v.sha(decision_driver) != decision_spec.get("driver_sha256"):
        v.fail("protocol.retention_decision", "retention decision driver binding is stale")
    decision = module(decision_name)
    decision_path_name = bundle_json_name(v, decision_spec.get("path"), "retention decision path")
    retained_decision = v.load(v.bundle_path(decision_path_name, "retention decision"))
    derived_decision = decision.derive(ROOT, ROOT / summary_path_name,
        decision_spec.get("large_peak_reduction_threshold_percent"))
    if decision.canonical(retained_decision) != decision.canonical(derived_decision):
        v.fail(decision_path_name, "retained decision differs from independent derivation")

    return {
        "status": "pass",
        "change": CHANGE,
        "stage": stage,
        "terminal_receipts": terminal_receipts,
        "pilots": pilots,
        "preparatory_profiles": preparatory_profiles,
        "formal_profiles": formal_profiles,
        "summary_rederived": True,
        "preparatory_summary_rederived": True,
        "retention_decision_rederived": True,
        "formal_reports": matrix["formal_reports"],
        "retained_samples": matrix["retained_samples"],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--stage", choices=("precleanup", "aftercleanup", "final"), default="final")
    args = parser.parse_args()
    try:
        print(json.dumps(validate(args.stage), sort_keys=True))
    except (OSError, ValueError, KeyError, TypeError, AssertionError) as error:
        print("INVALID: " + str(error), file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
