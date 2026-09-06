#!/usr/bin/env python3
"""Validate the terminal 0438 ODP evidence lifecycle.

This driver is read-only.  It delegates formal report/profile validation to
the retained ``verify.py`` module, independently rederives ``summary.json``,
checks the two-role pilot and profile custody, and binds the historical
0437-derived hypothesis without counting it as a 0438 profile.  Cleanup and
candidate-retention decisions are protocol inputs; names, counts, or hashes
are never inferred from the live checkout.

The capture protocol is immutable after formal receipts bind its hash.  The
separate ``lifecycle-contract.json`` sidecar therefore supplies
``precleanup_receipt``, ``cleanup_receipt``, ``cleanup_inventory``,
``preserved_paths``, ``goal_path``, ``goal_sha256``, ``summary_path`` /
``summary_module``, and a structured ``retention_gate`` with a decision
driver/path and retained candidate source/patch artifacts.  A draft sidecar
without a structured gate is accepted only at ``--stage precleanup``; final
validation fails closed until those bindings exist.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
from pathlib import Path
import re
from typing import Any


ROOT = Path(__file__).resolve().parent
CHANGE = 438
ROLES = ("before-streaming", "after-streaming")
MODES = ("normal", "allocator")
SHAPES = {"tiny": 64, "medium": 4_096, "large": 8_192}
PROFILE_KINDS = ("stat", "record")
HEX40 = re.compile(r"^[0-9a-fA-F]{40}$")
HEX64 = re.compile(r"^[0-9a-fA-F]{64}$")
STAGES = ("precleanup", "aftercleanup", "final")


class LifecycleError(ValueError):
    pass


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


def canonical(value: Any) -> bytes:
    try:
        return json.dumps(
            value,
            sort_keys=True,
            separators=(",", ":"),
            ensure_ascii=False,
            allow_nan=False,
        ).encode("utf-8")
    except (TypeError, ValueError, OverflowError) as error:
        fail("canonical JSON", str(error))
    raise AssertionError("unreachable")


def safe_relative(value: Any, label: str, *, suffix: str | None = None) -> str:
    if not isinstance(value, str) or not value:
        fail(label, "expected a non-empty bundle-relative path")
    path = Path(value)
    if path.is_absolute() or not path.parts or ".." in path.parts or path.parts[0] in {"", "."}:
        fail(label, "path is absolute or escapes the bundle")
    if suffix is not None and path.suffix != suffix:
        fail(label, f"path must end with {suffix}")
    return value


def bundle_path(value: Any, label: str, *, suffix: str | None = None) -> Path:
    name = safe_relative(value, label, suffix=suffix)
    path = (ROOT / name).resolve()
    if not path.is_relative_to(ROOT.resolve()):
        fail(label, "path escapes the bundle")
    return path


def module_stem(value: Any, label: str) -> str:
    name = safe_relative(value, label, suffix=".py")
    path = Path(name)
    if len(path.parts) != 1:
        fail(label, "module must be a single bundle-level filename")
    return path.stem


def load_module(name: Any, label: str):
    stem = module_stem(name, label)
    path = ROOT / f"{stem}.py"
    if not path.is_file():
        fail(label, f"module is missing: {path.name}")
    spec = importlib.util.spec_from_file_location(f"change0438_lifecycle_{stem}", path)
    if spec is None or spec.loader is None:
        fail(label, f"cannot load {path.name}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def load_protocol() -> dict[str, Any]:
    value = obj(load_json(ROOT / "protocol.json", "protocol.json"), "protocol.json")
    if value.get("change") != CHANGE:
        fail("protocol.change", f"expected {CHANGE}")
    return value


def load_lifecycle_contract(protocol: dict[str, Any]) -> dict[str, Any]:
    """Load lifecycle-only bindings without changing the capture protocol."""
    path = ROOT / "lifecycle-contract.json"
    value = obj(load_json(path, "lifecycle-contract.json"), "lifecycle-contract.json")
    if value.get("schema") != "litchi-0438-lifecycle-contract-v1" or value.get("change") != CHANGE:
        fail("lifecycle-contract.json", "schema or change differs")
    if value.get("protocol_path") != "protocol.json" or value.get("protocol_sha256") != sha(ROOT / "protocol.json"):
        fail("lifecycle-contract.json", "capture protocol binding is stale")
    return value


def verify_module():
    return load_module("verify.py", "protocol.verifier")


def required_paths(contract: Any, key: str, count: int, label: str) -> tuple[str, ...]:
    value = obj(contract, label).get(key)
    if not isinstance(value, list) or len(value) != count:
        fail(f"{label}.{key}", f"expected exactly {count} paths")
    result = tuple(safe_relative(item, f"{label}.{key}") for item in value)
    if len(set(result)) != len(result):
        fail(f"{label}.{key}", "contains duplicate paths")
    return result


def protocol_contract(value: dict[str, Any]) -> dict[str, Any]:
    if value.get("samples") != 30 or value.get("warmups") != 3 or value.get("repeats") != 2:
        fail("protocol", "samples/warmups/repeats differ from 30/3/2")
    if value.get("modes") != list(MODES) or value.get("shapes") != SHAPES:
        fail("protocol", "mode or shape matrix differs from 64/4096/8192")
    roles = obj(value.get("roles"), "protocol.roles")
    if tuple(roles) != ROLES:
        fail("protocol.roles", "must contain ordered before-streaming and after-streaming roles")
    for role in ROLES:
        spec = obj(roles.get(role), f"protocol.roles.{role}")
        if spec.get("selector") != "odp_streaming_create" or spec.get("source_field") != "odp_slides" or spec.get("source_role") != "streaming":
            fail(f"protocol.roles.{role}", "selector/source binding differs")
        if spec.get("oracle_role") != "after-streaming":
            fail(f"protocol.roles.{role}.oracle_role", "must use the copied after-streaming oracle")
        if spec.get("build_directory") not in {"before", "after"}:
            fail(f"protocol.roles.{role}.build_directory", "must be before or after")

    order = value.get("order")
    if not isinstance(order, list) or len(order) != 12:
        fail("protocol.order", "must contain the 12 fixed lanes")
    expected_order = [
        ("normal", "tiny", "R1"),
        ("normal", "medium", "R1"),
        ("normal", "large", "R1"),
        ("allocator", "tiny", "R1"),
        ("allocator", "medium", "R1"),
        ("allocator", "large", "R1"),
        ("allocator", "large", "R2"),
        ("allocator", "medium", "R2"),
        ("allocator", "tiny", "R2"),
        ("normal", "large", "R2"),
        ("normal", "medium", "R2"),
        ("normal", "tiny", "R2"),
    ]
    actual_order = []
    for index, item in enumerate(order):
        lane = obj(item, f"protocol.order[{index}]")
        actual_order.append((lane.get("mode"), lane.get("shape"), lane.get("repeat")))
    if actual_order != expected_order:
        fail("protocol.order", "lane order differs from frozen A1/B1/B2/A2 contract")

    matrix = obj(value.get("matrix"), "protocol.matrix")
    if tuple(matrix.get("roles", ())) != ROLES or matrix.get("formal_reports") != 24 or matrix.get("retained_samples") != 720:
        fail("protocol.matrix", "must bind two roles, 24 reports, and 720 samples")
    report_contract = obj(value.get("report_contract"), "protocol.report_contract")
    if report_contract.get("required_reports") != 24 or report_contract.get("retained_samples") != 720:
        fail("protocol.report_contract", "must bind 24 reports and 720 samples")

    oracle = obj(value.get("oracle"), "protocol.oracle")
    oracle_path = bundle_path(oracle.get("path"), "protocol.oracle.path")
    verifier_path = bundle_path(oracle.get("verifier_path"), "protocol.oracle.verifier_path")
    if oracle.get("sha256") != sha(oracle_path) or oracle.get("verifier_sha256") != sha(verifier_path):
        fail("protocol.oracle", "retained oracle hashes are stale")

    pilots = required_paths(value.get("pilot_contract"), "required_receipts", 12, "protocol.pilot_contract")
    profiles = required_paths(value.get("formal_profiles"), "required_receipts", 4, "protocol.formal_profiles")
    return {"pilots": pilots, "profiles": profiles}


def validate_source_custody(v: Any, path: Path, row: dict[str, Any]) -> None:
    before = row.get("source_before")
    after = row.get("source_after")
    if before is None or after is None:
        fail(str(path), "source_before/source_after are required")
    v.check_source_manifest(before, f"{path}.source_before")
    v.check_source_manifest(after, f"{path}.source_after")
    if before != after or row.get("source_unchanged") is not True:
        fail(str(path), "source custody is not unchanged")


def validate_receipt_artifacts(v: Any, path: Path, row: dict[str, Any]) -> None:
    artifacts = row.get("artifacts")
    if not isinstance(artifacts, dict):
        fail(str(path), "artifact inventory is missing")
    for name, record in artifacts.items():
        if not isinstance(name, str):
            fail(str(path), "artifact name is not a string")
        v.artifact_data(path, {**record, "path": name}, f"{path}.artifacts.{name}")
    log = row.get("log")
    if log is not None:
        v.artifact_data(path, log, f"{path}.log")


def build_descriptors(v: Any, protocol_sha: str) -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    for role, directory in (("before-streaming", "before"), ("after-streaming", "after")):
        path = ROOT / directory / "build.json"
        row = obj(v.load(path, str(path)), str(path))
        if row.get("schema") != "litchi-0438-build-descriptor-v1" or row.get("change") != CHANGE:
            fail(str(path), "build descriptor identity differs")
        if row.get("protocol_sha256") != protocol_sha:
            fail(str(path), "build descriptor protocol binding is stale")
        source = obj(row.get("source_manifest"), f"{path}.source_manifest")
        v.check_source_manifest(source, f"{path}.source_manifest")
        build_receipt = v.bundle_path(row.get("build_receipt"), f"{path}.build_receipt")
        copies_receipt = v.bundle_path(row.get("binary_copies_receipt"), f"{path}.binary_copies_receipt")
        if row.get("build_receipt_sha256") != v.sha(build_receipt) or row.get("binary_copies_receipt_sha256") != v.sha(copies_receipt):
            fail(str(path), "build/copy receipt hash is stale")
        binaries = obj(row.get("binaries"), f"{path}.binaries")
        copies = obj(v.load(copies_receipt, str(copies_receipt)), str(copies_receipt))
        if set(binaries) != set(MODES) or binaries != copies:
            fail(str(path), "binary descriptor differs from retained copies")
        result[role] = row
    return result


def validate_pilots(
    v: Any,
    contract: tuple[str, ...],
    builds: dict[str, dict[str, Any]],
    protocol: dict[str, Any],
) -> int:
    oracle = obj(protocol.get("oracle"), "protocol.oracle")
    oracle_name = oracle.get("verifier_path")
    oracle_protocol = oracle.get("path")
    oracle_module = load_module(oracle_name, "protocol.oracle.verifier_path")
    for name in contract:
        path = v.bundle_path(name, "pilot receipt")
        row = obj(v.load(path, name), name)
        parts = Path(name).parts
        if len(parts) != 4 or parts[0] != "pilots" or parts[2] != "initial" or parts[1] not in ROLES:
            fail(name, "pilot path does not match the frozen grammar")
        stem = Path(parts[3]).stem
        expected = stem.removesuffix("-receipt")
        if "-" not in expected:
            fail(name, "pilot filename has no mode/shape")
        mode, shape = expected.rsplit("-", 1)
        if mode not in MODES or shape not in SHAPES or row.get("role") != parts[1] or row.get("mode") != mode or row.get("shape") != shape:
            fail(name, "pilot path/role/mode/shape binding differs")
        if row.get("change") != CHANGE or row.get("phase") != "preparatory-pilot" or row.get("status") != "pass" or row.get("exit_code") != 0 or row.get("oracle_exit_code") != 0:
            fail(name, "pilot is not a passing terminal receipt")
        if row.get("samples") != 3 or row.get("warmups") != 1:
            fail(name, "pilot dimensions differ from 3 samples/1 warmup")
        validate_source_custody(v, path, row)
        if row.get("driver_sha256") != v.sha(ROOT / "pilot.py"):
            fail(name, "pilot driver custody is stale")
        if row.get("oracle_path") != oracle_name or row.get("protocol_path") != oracle_protocol:
            fail(name, "pilot oracle/protocol path binding differs")
        oracle_path = v.bundle_path(oracle_name, f"{name}.oracle_path")
        oracle_protocol_path = v.bundle_path(oracle_protocol, f"{name}.protocol_path")
        if row.get("oracle_sha256") != v.sha(oracle_path) or row.get("protocol_sha256") != v.sha(oracle_protocol_path):
            fail(name, "pilot oracle/protocol hash binding is stale")
        build_dir = "before" if parts[1] == "before-streaming" else "after"
        build_path = v.bundle_path(row.get("build_receipt"), f"{name}.build_receipt")
        if row.get("build_receipt_sha256") != v.sha(build_path):
            fail(name, "pilot build receipt hash is stale")
        build = builds[parts[1]]
        binaries = obj(build.get("binaries"), f"{build_dir}/build.json.binaries")
        if row.get("source_before") != build.get("source_manifest"):
            fail(name, "pilot source manifest differs from its role build")
        if row.get("binary") != binaries.get(mode):
            fail(name, "pilot binary differs from its role descriptor")
        validate_receipt_artifacts(v, path, row)
        report_name = f"pilots/{parts[1]}/initial/{mode}-{shape}.json"
        report_path = bundle_path(report_name, f"{name}.report")
        try:
            oracle_module.validate_report(
                report_path,
                mode,
                shape,
                "after-streaming",
                samples=3,
                warmups=1,
            )
        except (OSError, KeyError, TypeError, ValueError, AssertionError) as error:
            fail(name, f"pilot report/oracle validation failed: {error}")
    return len(contract)


def validate_formal_profile_paths(v: Any, names: tuple[str, ...]) -> int:
    for name in names:
        parts = Path(name).parts
        if len(parts) != 4 or parts[0] != "profiles" or parts[1] not in ROLES or parts[2] not in PROFILE_KINDS or parts[3] != "receipt.json":
            fail(name, "formal profile path does not match the frozen grammar")
        if "preparatory" in parts:
            fail(name, "historical/preparatory profile cannot enter formal count")
        path = v.bundle_path(name, "formal profile")
        row = obj(v.load(path, name), name)
        if row.get("status") != "pass" or row.get("change") != CHANGE or row.get("role") != parts[1] or row.get("kind") != parts[2] or row.get("attempt") != "formal" or row.get("preparatory") is not False or row.get("shape") != "large":
            fail(name, "formal profile receipt fields differ from its path")
    return len(names)


def validate_prior_profile(
    v: Any, builds: dict[str, dict[str, Any]], contract: dict[str, Any]
) -> dict[str, Any]:
    path = ROOT / "prior-profile" / "reference.json"
    reference = obj(v.load(path, str(path)), str(path))
    if reference.get("change") != CHANGE:
        fail(str(path), "historical reference must be labeled change 438")
    scope = reference.get("scope")
    if not isinstance(scope, str) or "hypothesis" not in scope.lower() or "not" not in scope.lower() or "new complete" not in scope.lower():
        fail(str(path), "reference must state historical hypothesis-only scope")
    before_manifest = obj(builds["before-streaming"].get("source_manifest"), "before/build.json.source_manifest")
    if reference.get("source_manifest_sha256") != before_manifest.get("sha256"):
        fail(str(path), "historical profile source manifest differs from fresh before build")
    profiles = obj(reference.get("profiles"), f"{path}.profiles")
    if set(profiles) != set(PROFILE_KINDS):
        fail(str(path), "historical reference must contain stat and record only")
    hypothesis_module = load_module(
        contract.get("hypothesis_module", "before-hypothesis.py"),
        "historical hypothesis module",
    )
    hypothesis_name = safe_relative(
        contract.get("hypothesis_path", "before-hypothesis.json"),
        "lifecycle.hypothesis_path",
        suffix=".json",
    )
    hypothesis_path = bundle_path(hypothesis_name, "historical hypothesis summary")
    derive_hypothesis = getattr(hypothesis_module, "derive", None)
    if not callable(derive_hypothesis):
        fail("before-hypothesis.py", "historical hypothesis module must expose derive")
    if canonical(v.load(hypothesis_path, hypothesis_name)) != canonical(derive_hypothesis()):
        fail(hypothesis_name, "historical hypothesis summary differs from derivation")
    before_binary = obj(obj(builds["before-streaming"].get("binaries"), "before binaries").get("normal"), "before normal binary")
    retained: dict[str, Any] = {}
    for kind in PROFILE_KINDS:
        item = obj(profiles.get(kind), f"{path}.profiles.{kind}")
        nested = obj(item.get("receipt"), f"{path}.profiles.{kind}.receipt")
        if nested.get("change") != 437 or nested.get("status") != "pass" or nested.get("role") != "after-streaming" or nested.get("kind") != kind:
            fail(f"{path}.profiles.{kind}", "nested receipt is not the retained 0437 after-streaming profile")
        nested_source = obj(nested.get("source_manifest"), f"{path}.profiles.{kind}.source_manifest")
        if nested_source.get("sha256") != before_manifest.get("sha256"):
            fail(f"{path}.profiles.{kind}", "nested source manifest differs from fresh before source")
        binary = obj(nested.get("binary"), f"{path}.profiles.{kind}.binary")
        if binary.get("sha256") != before_binary.get("sha256") or binary.get("bytes") != before_binary.get("bytes"):
            fail(f"{path}.profiles.{kind}", "historical executable does not match fresh before binary")
        retained_name = safe_relative(item.get("retained_hotspot_artifact"), f"{path}.profiles.{kind}.retained_hotspot_artifact")
        artifact_key = "perf_stat" if kind == "stat" else "perf_report"
        original = nested["artifacts"][artifact_key]
        _, logical = v.artifact_data(path, {**original, "path": retained_name},
                                     f"{path}.profiles.{kind}.retained_hotspot_artifact")
        retained[kind] = {"receipt": nested, "artifact": retained_name,
                          "artifact_sha256": hashlib.sha256(logical).hexdigest()}
    return {"formal_profiles": 0, "historical_profiles": 2, "retained": retained}


def validate_cleanup(v: Any, stage: str, contract: dict[str, Any]) -> dict[str, Any]:
    if stage == "precleanup":
        return {"stage": stage, "checked": False}
    required = ("precleanup_receipt", "cleanup_receipt", "cleanup_inventory", "preserved_paths", "goal_path", "goal_sha256")
    if any(contract.get(key) in (None, "", []) for key in required):
        fail("protocol.cleanup", "final lifecycle bindings are not frozen")
    pre_name = safe_relative(contract["precleanup_receipt"], "lifecycle.precleanup_receipt", suffix=".json")
    cleanup_name = safe_relative(contract["cleanup_receipt"], "lifecycle.cleanup_receipt", suffix=".json")
    inventory_name = safe_relative(contract["cleanup_inventory"], "lifecycle.cleanup_inventory", suffix=".json")
    for name in (pre_name, cleanup_name):
        row = obj(v.load(bundle_path(name, name), name), name)
        if row.get("status") != "pass" or row.get("exit_code") != 0:
            fail(name, "cleanup lifecycle receipt is not passing")
    cleanup_paths = contract.get("cleanup_paths")
    if cleanup_paths != contract.get("cleanup_expected_paths") or not isinstance(cleanup_paths, list):
        fail("lifecycle.cleanup_paths", "cleanup allowlist is not explicit and equal")
    inventory = obj(v.load(bundle_path(inventory_name, inventory_name), inventory_name), inventory_name)
    if inventory.get("change") != CHANGE or inventory.get("status") != "pass" or inventory.get("protocol") != "protocol.json":
        fail(inventory_name, "cleanup inventory is not a passing 0438 record")
    if inventory.get("cleanup_paths") != cleanup_paths or inventory.get("cleanup_expected_paths") != cleanup_paths:
        fail(inventory_name, "cleanup inventory path binding differs")
    removed = inventory.get("removed")
    if not isinstance(removed, list) or len(removed) != len(cleanup_paths):
        fail(inventory_name, "cleanup inventory removed list differs")
    for expected, row in zip(cleanup_paths, removed):
        if not isinstance(row, dict) or row.get("path") != expected:
            fail(inventory_name, "cleanup inventory removed path differs")
        for key in ("regular_files", "regular_file_bytes"):
            if isinstance(row.get(key), bool) or not isinstance(row.get(key), int) or row[key] < 0:
                fail(inventory_name, f"invalid removed {key}")
    preserved = contract.get("preserved_paths")
    if not isinstance(preserved, list) or not preserved or set(inventory.get("preserved_target_directory_identity", {})) != set(preserved):
        fail(inventory_name, "preserved target identities differ")
    if inventory.get("user_goal_path") != contract.get("goal_path") or inventory.get("user_goal_sha256") != contract.get("goal_sha256"):
        fail(inventory_name, "GOAL custody differs")
    return {"stage": stage, "checked": True, "inventory": inventory_name}


def validate_retention(v: Any, stage: str, contract: dict[str, Any], summary_name: str) -> dict[str, Any]:
    gate = contract.get("retention_gate")
    if not isinstance(gate, dict):
        if stage == "precleanup":
            return {"status": "pending", "rederived": False}
        fail("protocol.retention_gate", "must be expanded to a structured final decision binding")
    if (
        gate.get("minimum_p50_improvement_percent") != 5.0
        or gate.get("required_shapes") != ["medium", "large"]
        or gate.get("required_repeats") != ["R1", "R2"]
        or gate.get("require_mean_ci_nonoverlap") is not True
        or gate.get("require_allocation_nonincrease") is not True
    ):
        fail("protocol.retention_gate", "acceptance must require >=5% p50 gain for medium/large R1/R2, non-overlapping mean CIs, and no allocation increase")
    policy = gate.get("rejection_policy")
    if not isinstance(policy, str) or "retain" not in policy.lower() or "source" not in policy.lower() or "patch" not in policy.lower():
        fail("protocol.retention_gate.rejection_policy", "must preserve candidate source and patch on rejection")
    artifacts = gate.get("candidate_artifacts")
    if not isinstance(artifacts, list) or not artifacts:
        fail("protocol.retention_gate.candidate_artifacts", "candidate source/patch artifact list is required")
    for index, item in enumerate(artifacts):
        if isinstance(item, str):
            name = safe_relative(item, f"protocol.retention_gate.candidate_artifacts[{index}]")
            path = bundle_path(name, f"protocol.retention_gate.candidate_artifacts[{index}]")
        else:
            record = obj(item, f"protocol.retention_gate.candidate_artifacts[{index}]")
            name = safe_relative(record.get("path"), f"protocol.retention_gate.candidate_artifacts[{index}].path")
            path = bundle_path(name, f"protocol.retention_gate.candidate_artifacts[{index}].path")
            if record.get("sha256") != sha(path):
                fail(f"protocol.retention_gate.candidate_artifacts[{index}]", "candidate artifact hash is stale")
        if not path.is_file():
            fail(name, "candidate retention artifact is missing")
    decision = obj(gate.get("decision"), "protocol.retention_gate.decision")
    driver_name = module_stem(decision.get("driver"), "protocol.retention_gate.decision.driver")
    driver = ROOT / f"{driver_name}.py"
    if decision.get("driver_sha256") != sha(driver):
        fail("protocol.retention_gate.decision.driver_sha256", "decision driver hash is stale")
    decision_name = safe_relative(decision.get("path"), "protocol.retention_gate.decision.path", suffix=".json")
    retained = obj(v.load(bundle_path(decision_name, decision_name), decision_name), decision_name)
    decision_module = load_module(f"{driver_name}.py", "protocol.retention_gate.decision.driver")
    derive = getattr(decision_module, "derive", None)
    if not callable(derive):
        fail("protocol.retention_gate.decision.driver", "decision module must expose derive")
    derived = derive(ROOT, ROOT / summary_name, gate)
    if canonical(retained) != canonical(derived):
        fail(decision_name, "retained decision differs from independent derivation")
    if retained.get("status") not in {"accepted", "rejected"}:
        fail(decision_name, "decision status must be accepted or rejected")
    return {"status": retained["status"], "rederived": True, "decision": decision_name}


def validate(stage: str) -> dict[str, Any]:
    protocol = load_protocol()
    lifecycle = load_lifecycle_contract(protocol)
    contract = protocol_contract(protocol)
    v = verify_module()
    protocol_sha = sha(ROOT / "protocol.json")
    builds = build_descriptors(v, protocol_sha)
    # Formal verifier owns report, source, binary, phase, profile, identity,
    # and allocator-vector checks.  Prior-profile artifacts are never counted.
    envelope = v.verify_matrix(require_binaries=False)
    if envelope.get("matrix", {}).get("formal_reports") != 24 or envelope.get("matrix", {}).get("retained_samples") != 720:
        fail("verify.py", "formal matrix dimensions differ from protocol")
    pilots = validate_pilots(v, contract["pilots"], builds, protocol)
    profiles = validate_formal_profile_paths(v, contract["profiles"])
    prior = validate_prior_profile(v, builds, lifecycle)
    summary_name = safe_relative(lifecycle.get("summary_path", "summary.json"), "lifecycle.summary_path", suffix=".json")
    summary_module = load_module(lifecycle.get("summary_module", "summary.py"), "lifecycle.summary_module")
    retained_summary = v.load(bundle_path(summary_name, "summary_path"), summary_name)
    derive = getattr(summary_module, "derive", None)
    if not callable(derive):
        fail("protocol.summary_module", "summary module must expose derive")
    if canonical(retained_summary) != canonical(derive()):
        fail(summary_name, "retained summary differs from independent derivation")
    cleanup = validate_cleanup(v, stage, lifecycle)
    retention = validate_retention(v, stage, lifecycle, summary_name)
    archived = v.load(ROOT / "candidate-artifacts.json", "candidate artifacts")
    for item in archived:
        v.artifact_data(ROOT / "candidate-artifacts.json", item, "candidate artifact")
    for role, snapshot in (("before-streaming", "before-streaming.rs.txt"),
                           ("after-streaming", "after-streaming.rs.txt")):
        sources = v.load(ROOT / builds[role]["source_manifest"]["path"], "candidate source manifest")
        if sha(ROOT / "candidate" / snapshot) != sources["crates/litchi-odp/src/streaming.rs"]:
            fail(snapshot, "archived candidate source differs from its build")
    if sha(ROOT / "candidate/markup_tests.rs.txt") != sources["crates/litchi-odp/src/streaming/markup_tests.rs"]:
        fail("candidate tests", "archived tests differ from candidate build")
    expected_path = ROOT / "expected-checks.json"
    terminal_count = 0
    if expected_path.exists():
        expected = v.load(expected_path, "expected command receipts")
        actual = {}
        for path in ROOT.rglob("*.json"):
            row = v.load(path, str(path))
            if not isinstance(row, dict) or "source_before" not in row:
                continue
            name = path.relative_to(ROOT).as_posix()
            if row.get("status") not in {"pass", "failed"} or row.get("source_before") != row.get("source_after"):
                fail(name, "nonterminal receipt or changed source custody")
            v.check_source_manifest(row["source_before"], name)
            if path.parent.name == "checks":
                if row.get("driver_sha256") != v.sha(ROOT / "check.py"):
                    fail(name, "command custody driver differs")
                if "log" in row:
                    v.artifact_data(path, row["log"], name + ".log")
            actual[path.relative_to(ROOT).with_suffix("").as_posix()] = row["status"]
        if actual != expected:
            fail("expected-checks.json", "terminal source-command receipt set differs")
        terminal_count = len(actual)
    return {
        "status": "pass",
        "change": CHANGE,
        "stage": stage,
        "formal_reports": 24,
        "retained_samples": 720,
        "pilots": pilots,
        "formal_profiles": profiles,
        "historical_prior_profiles": prior["historical_profiles"],
        "formal_summary_rederived": True,
        "cleanup": cleanup,
        "retention": retention,
        "terminal_source_receipts": terminal_count,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--stage", choices=STAGES, default="final")
    args = parser.parse_args()
    try:
        print(json.dumps(validate(args.stage), sort_keys=True))
    except (OSError, KeyError, TypeError, ValueError, AssertionError) as error:
        print(f"INVALID: {error}")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
