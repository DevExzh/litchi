#!/usr/bin/env python3
"""Authenticate the retained 0457 evidence bundle.

The candidate, control, and native directories own their detailed schemas and
oracles.  This small outer verifier checks the shared check.py custody,
release-gate selection, derived-result bindings, fuzz inventories, and the
final bundle seal, then delegates the detailed replay to those verifiers.

The verifier is read-only.  ``--precleanup`` asks the child verifiers to check
the original temporary binaries and fixtures.  ``--portable`` copies the
sealed bundle to a temporary directory and runs the copied verifier without
the original checkout.  Missing or still-running late evidence is reported as
``pending`` (exit code 2), never as a pass.
"""

from __future__ import annotations

import argparse
import ast
import hashlib
import json
import re
import shutil
import subprocess
import sys
import tempfile
import zlib
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parent
SCHEMA = "litchi-0457-root-verification-v1"
CHANGE = 457
SHA256_RE = re.compile(r"[0-9a-f]{64}\Z")
RETRY_RE = re.compile(r"^(.*)-r([0-9]+)$")


class VerificationError(AssertionError):
    pass


def fail(message: str) -> None:
    raise VerificationError(message)


def load(path: Path, label: str) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"{label}: invalid JSON: {error}")
    raise AssertionError("unreachable")


def obj(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(f"{label}: expected object")
    return value


def text(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        fail(f"{label}: expected non-empty text")
    return value


def digest(value: Any, label: str) -> str:
    value = text(value, label).lower()
    if SHA256_RE.fullmatch(value) is None:
        fail(f"{label}: expected lowercase SHA-256")
    return value


def integer(value: Any, label: str, minimum: int = 0) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        fail(f"{label}: expected integer >= {minimum}")
    return value


def sha_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def regular(path: Path, label: str) -> Path:
    if not path.is_file() or path.is_symlink():
        fail(f"{label}: missing, symlink, or non-regular file: {path}")
    return path


def safe_path(base: Path, value: Any, label: str, *, must_exist: bool = True) -> Path:
    raw = text(value, label)
    relative = Path(raw)
    if relative.is_absolute() or ".." in relative.parts:
        fail(f"{label}: path must be relative and traversal-free")
    path = base / relative
    if any(part.is_symlink() for part in [base, *path.parents]):
        fail(f"{label}: path has a symlinked ancestor")
    try:
        resolved = path.resolve(strict=must_exist)
    except OSError as error:
        fail(f"{label}: cannot resolve path: {error}")
    if not resolved.is_relative_to(base.resolve()):
        fail(f"{label}: path escapes its base")
    if must_exist:
        regular(path, label)
    return path


def artifact(base: Path, value: Any, label: str) -> Path:
    row = obj(value, label)
    path = safe_path(base, row.get("path"), f"{label}.path")
    expected_bytes = integer(row.get("bytes"), f"{label}.bytes")
    expected_sha = digest(row.get("sha256"), f"{label}.sha256")
    if path.stat().st_size != expected_bytes or sha_file(path) != expected_sha:
        fail(f"{label}: artifact identity differs")
    return path


def source_record(value: Any, label: str) -> tuple[Path, dict[str, Any]]:
    row = obj(value, label)
    path = safe_path(ROOT, row.get("path"), f"{label}.path")
    expected_sha = digest(row.get("sha256"), f"{label}.sha256")
    expected_files = integer(row.get("files"), f"{label}.files")
    if sha_file(path) != expected_sha:
        fail(f"{label}: source manifest hash differs")
    manifest = obj(load(path, label), label)
    if len(manifest) != expected_files:
        fail(f"{label}: source manifest file count differs")
    return path, row


def check_log(receipt: dict[str, Any], label: str) -> None:
    log = obj(receipt.get("log"), f"{label}.log")
    path = artifact(ROOT, log, f"{label}.log")
    del path


def parent_revision() -> str:
    parent = obj(load(ROOT / "parent.json", "parent.json"), "parent.json")
    if parent.get("change") != CHANGE:
        fail("parent.json change differs")
    return text(parent.get("parent_revision"), "parent.json.parent_revision")


def check_receipt(path: Path, revision: str) -> dict[str, Any]:
    receipt = obj(load(path, str(path)), str(path))
    label = str(path.relative_to(ROOT))
    if receipt.get("change") != CHANGE:
        fail(f"{label}: receipt belongs to another change")
    if receipt.get("revision") != revision:
        fail(f"{label}: source revision differs from parent binding")
    if digest(receipt.get("driver_sha256"), f"{label}.driver_sha256") != sha_file(ROOT / "check.py"):
        fail(f"{label}: driver hash differs from check.py")
    cwd = text(receipt.get("cwd"), f"{label}.cwd")
    cwd_path = Path(cwd)
    if not cwd_path.is_absolute() or cwd_path != cwd_path.resolve():
        fail(f"{label}: cwd is not an absolute canonical path")
    status = receipt.get("status")
    if status not in {"running", "pass", "failed"}:
        fail(f"{label}: unsupported receipt status {status!r}")
    before = obj(receipt.get("source_before"), f"{label}.source_before")
    source_record(before, f"{label}.source_before")
    if status == "running":
        # A live command may not have emitted source_after, exit_code, or a
        # final log yet. Its source_before identity is still checked, and
        # audit_check_receipts reports it as pending.
        return receipt
    after = obj(receipt.get("source_after"), f"{label}.source_after")
    source_record(after, f"{label}.source_after")
    if before != after or receipt.get("source_unchanged") is not True:
        fail(f"{label}: source custody is not unchanged")
    exit_code = receipt.get("exit_code")
    if not isinstance(exit_code, int) or isinstance(exit_code, bool):
        fail(f"{label}: completed receipt has no integer exit code")
    if status == "pass" and exit_code != 0:
        fail(f"{label}: pass receipt has nonzero exit code")
    if status == "failed" and exit_code == 0:
        fail(f"{label}: failed receipt has zero exit code")
    check_log(receipt, label)
    return receipt


def audit_check_receipts(revision: str) -> tuple[dict[str, dict[str, Any]], list[str], list[str]]:
    receipts: dict[str, dict[str, Any]] = {}
    pending: list[str] = []
    failures: list[str] = []
    for path in sorted((ROOT / "checks").glob("*.json")):
        receipt = check_receipt(path, revision)
        name = path.stem
        receipts[name] = receipt
        if receipt.get("status") == "running":
            pending.append(f"checks/{path.name}: still running")
        # Failed receipts are immutable historical attempts. They remain
        # visible in the result and are authenticated above, but only a
        # required gate with no passing replacement makes the bundle fail.
    return receipts, pending, failures


def command_table() -> dict[str, list[str]]:
    """Read only the literal COMMANDS mapping; never execute run-checks.py."""

    tree = ast.parse((ROOT / "run-checks.py").read_text(encoding="utf-8"))
    for node in tree.body:
        if isinstance(node, ast.Assign) and any(
            isinstance(target, ast.Name) and target.id == "COMMANDS" for target in node.targets
        ):
            value = ast.literal_eval(node.value)
            if isinstance(value, dict) and all(isinstance(k, str) and isinstance(v, list) for k, v in value.items()):
                return value
    fail("run-checks.py has no literal COMMANDS mapping")
    raise AssertionError("unreachable")


REQUIRED_GATE_GROUPS: dict[str, tuple[str, ...]] = {
    "baseline-build": ("baseline-build",),
    "candidate-build": ("candidate-build",),
    "candidate-final-build": ("candidate-final-build",),
    "format-non-iwork": ("format-non-iwork",),
    "harness-format": ("harness-format",),
    "zip": ("zip",),
    "odf-common": ("odf-common",),
    "odp": ("odp",),
    "odp-append-final": ("odp-append-final",),
    "opc": ("opc",),
    "harness": ("harness",),
    "insertion-tests": ("insertion-tests",),
    "strict": ("strict",),
    "harness-strict": ("harness-strict",),
    "doc": ("doc",),
    "workspace": ("workspace",),
    "boundaries": ("boundaries",),
    "fixture-export": ("fixture-export",),
    "fixture-export-final": ("fixture-export-final",),
    "output-binding": ("output-binding",),
    "output-binding-final": ("output-binding-final",),
    "native-record-preflight": ("native-record-preflight",),
    "native": ("native",),
    "native-final-preflight": ("native-final-preflight",),
    "native-final": ("native-final",),
    "comparison": ("comparison",),
    "retained-plan-tests": ("retained-plan-tests",),
    "odf-common-final": ("odf-common-final",),
    "odp-retained-plan": ("odp-retained-plan",),
    "strict-retained-plan": ("strict-retained-plan",),
    "doc-retained-plan": ("doc-retained-plan",),
    "harness-strict-retained-plan": ("harness-strict-retained-plan",),
    "format-retained-plan": ("format-retained-plan",),
    "candidate-final-precleanup": ("candidate-final-precleanup",),
    "candidate-final-derive": ("candidate-final-derive",),
    "comparison-final": ("comparison-final",),
}


FINAL_SOURCE_GATE_NAMES = (
    "candidate-final-build",
    "fixture-export-final",
    "output-binding-final",
    "native-final-preflight",
    "native-final",
    "strict-retained-plan-r1",
    "odp-retained-plan",
    "doc-retained-plan",
    "harness-strict-retained-plan",
    "format-retained-plan",
    "candidate-final-precleanup",
    "candidate-final-derive",
    "comparison-final",
)

FINAL_TEST_COUNTS = {
    "retained-plan-tests": (13, 0, 0),
    "odf-common-final": (499, 0, 1),
    "odp-retained-plan": (368, 0, 0),
}


def gate_name_matches(name: str, base: str) -> bool:
    return name == base or bool(re.fullmatch(re.escape(base) + r"-r[0-9]+", name))


def gate_base(name: str) -> str:
    match = RETRY_RE.fullmatch(name)
    return match.group(1) if match else name


def select_gate(group: str, receipts: dict[str, dict[str, Any]], pending: list[str], failures: list[str]) -> str | None:
    names = [name for name in receipts if any(gate_name_matches(name, base) for base in REQUIRED_GATE_GROUPS[group])]
    passing = [name for name in names if receipts[name].get("status") == "pass"]
    if passing:
        def order(name: str) -> tuple[int, str]:
            match = RETRY_RE.fullmatch(name)
            return (int(match.group(2)) if match else 0, name)
        return max(passing, key=order)
    running = [name for name in names if receipts[name].get("status") == "running"]
    if running:
        pending.append(f"required gate {group} is still running: {', '.join(sorted(running))}")
    elif names:
        failures.append(f"required gate {group} has no passing receipt")
    else:
        pending.append(f"required gate {group} has not been captured")
    return None


def check_required_gates(receipts: dict[str, dict[str, Any]], pending: list[str], failures: list[str]) -> dict[str, str]:
    commands = command_table()
    selected: dict[str, str] = {}
    for group in REQUIRED_GATE_GROUPS:
        name = select_gate(group, receipts, pending, failures)
        if name is None:
            continue
        selected[group] = name
        base = gate_base(name)
        expected = commands.get(base)
        if expected is not None and receipts[name].get("argv") != expected:
            failures.append(f"checks/{name}.json: argv differs from run-checks.py COMMANDS[{base!r}]")
    return selected


def check_final_gate_counts(receipts: dict[str, dict[str, Any]], failures: list[str]) -> None:
    """Check the retained test totals for the final source epoch gates."""

    for name, expected in FINAL_TEST_COUNTS.items():
        receipt = receipts.get(name)
        if receipt is None or receipt.get("status") != "pass":
            continue
        observed = tuple(receipt.get(key) for key in ("passed_tests", "failed_tests", "ignored_tests"))
        if observed != expected:
            failures.append(f"checks/{name}.json: test totals differ (expected {expected}, got {observed})")


def check_build_bindings(pending: list[str], failures: list[str], precleanup: bool, revision: str) -> None:
    for role, receipt_name, subdir in (
        ("candidate", "candidate-build", "candidate"),
        ("control", "baseline-build", "control"),
    ):
        local = ROOT / subdir / "build-receipt.json"
        root_receipt = ROOT / "checks" / f"{receipt_name}.json"
        if not local.is_file() or not root_receipt.is_file():
            pending.append(f"{role} build binding is incomplete")
            continue
        if local.read_bytes() != root_receipt.read_bytes():
            failures.append(f"{role} build receipt differs from checks/{receipt_name}.json")
        build = obj(load(local, str(local)), str(local))
        if build.get("status") != "pass" or build.get("source_unchanged") is not True or build.get("revision") != revision:
            failures.append(f"{role} build receipt is not a successful parent-revision build")
    retention_path = ROOT / "candidate-retention.json"
    if not retention_path.is_file():
        pending.append("candidate-retention.json is missing")
        return
    retention = obj(load(retention_path, str(retention_path)), "candidate-retention.json")
    build_ref = obj(retention.get("build_receipt"), "candidate-retention.build_receipt")
    source_ref = obj(retention.get("source_manifest"), "candidate-retention.source_manifest")
    if build_ref.get("sha256") != sha_file(ROOT / "checks/candidate-build.json"):
        failures.append("candidate-retention build receipt hash differs")
    source_record(source_ref, "candidate-retention.source_manifest")
    if source_ref.get("sha256") != sha_file(ROOT / "candidate/source-manifest.json"):
        failures.append("candidate-retention source manifest hash differs")
    binaries = retention.get("binaries")
    if not isinstance(binaries, list) or len(binaries) != 3:
        failures.append("candidate-retention must bind normal, allocator, and fixture binaries")
    else:
        for index, raw in enumerate(binaries):
            row = obj(raw, f"candidate-retention.binaries[{index}]")
            source = Path(text(row.get("source"), f"candidate-retention.binaries[{index}].source"))
            if source.is_absolute() or ".." in source.parts:
                failures.append(f"candidate-retention binary {index} source path escapes repository")
            copy = Path(text(row.get("path"), f"candidate-retention.binaries[{index}].path"))
            if not copy.is_absolute() or not copy.is_relative_to(Path("/tmp/litchi-goal-0457")):
                failures.append(f"candidate-retention binary {index} copy path is not task-owned")
            expected_bytes = integer(row.get("bytes"), f"candidate-retention.binaries[{index}].bytes", 1)
            expected_sha = digest(row.get("sha256"), f"candidate-retention.binaries[{index}].sha256")
            if precleanup:
                # The final candidate build deliberately reuses the shared
                # Cargo target directory, so its source executable may have
                # replaced this historical epoch by the time root precleanup
                # runs.  The retained task-owned copy is the immutable
                # precleanup witness for this old candidate binding.
                if not copy.is_file() or copy.is_symlink() or copy.stat().st_size != expected_bytes or sha_file(copy) != expected_sha:
                    failures.append(f"candidate-retention binary {index} copy differs at precleanup")


def check_final_build_binding(
    receipts: dict[str, dict[str, Any]], pending: list[str], failures: list[str], revision: str
) -> dict[str, Any] | None:
    """Authenticate the final candidate build and return its source identity."""

    final_root = ROOT / "candidate-final"
    local_path = final_root / "build-receipt.json"
    gate_path = ROOT / "checks/candidate-final-build.json"
    if not local_path.is_file() or not gate_path.is_file():
        pending.append("candidate-final build binding is incomplete")
        return None
    try:
        if local_path.read_bytes() != gate_path.read_bytes():
            fail("candidate-final build receipt differs from checks/candidate-final-build.json")
        build = obj(load(local_path, str(local_path)), str(local_path))
        if build.get("change") != CHANGE or build.get("status") != "pass" or build.get("exit_code") != 0 or build.get("revision") != revision or build.get("source_unchanged") is not True:
            fail("candidate-final build receipt is not a successful parent-revision build")
        before = obj(build.get("source_before"), "candidate-final build source_before")
        after = obj(build.get("source_after"), "candidate-final build source_after")
        source_record(before, "candidate-final build source_before")
        source_record(after, "candidate-final build source_after")
        if before != after:
            fail("candidate-final build source custody differs")
        source_manifest = final_root / "source-manifest.json"
        if not source_manifest.is_file() or sha_file(source_manifest) != digest(after.get("sha256"), "candidate-final build source_after.sha256"):
            fail("candidate-final source manifest does not match its build receipt")
        binding_path = final_root / "binary-bindings.json"
        if not binding_path.is_file():
            pending.append("candidate-final binary binding is missing")
            return after
        binding = obj(load(binding_path, str(binding_path)), "candidate-final binary bindings")
        if binding.get("revision") != revision or binding.get("build_receipt") != {"path": "build-receipt.json", "sha256": sha_file(local_path)}:
            fail("candidate-final binary binding build identity differs")
        expected_source = {"path": "source-manifest.json", "sha256": sha_file(source_manifest), "files": integer(after.get("files"), "candidate-final source files")}
        if binding.get("source_manifest") != expected_source:
            fail("candidate-final binary binding source identity differs")
        binaries = obj(binding.get("binaries"), "candidate-final binary bindings.binaries")
        if set(binaries) != {"normal", "allocator"}:
            fail("candidate-final binary binding mode set differs")
        for mode, raw in binaries.items():
            row = obj(raw, f"candidate-final binary bindings.binaries.{mode}")
            if row.get("profile") != "release" or row.get("executable") is not True:
                fail(f"candidate-final {mode} binary binding is not a release executable")
            integer(row.get("bytes"), f"candidate-final {mode} binary bytes", 1)
            digest(row.get("sha256"), f"candidate-final {mode} binary sha256")
            source = Path(text(row.get("source_path"), f"candidate-final {mode} binary source_path"))
            if source.is_absolute() or ".." in source.parts:
                fail(f"candidate-final {mode} binary source path escapes the repository")
            copy = Path(text(row.get("copy_path"), f"candidate-final {mode} binary copy_path"))
            if not copy.is_absolute() or not copy.is_relative_to(Path("/tmp/litchi-goal-0457/candidate-final")):
                fail(f"candidate-final {mode} binary copy path is not task-owned")
        return after
    except (OSError, KeyError, TypeError, ValueError, VerificationError) as error:
        failures.append(str(error))
        return None


def final_capture_complete(phase: str, pending: list[str], failures: list[str]) -> bool:
    state_path = ROOT / "candidate-final" / "runs" / phase / "final" / "capture-state.json"
    if not state_path.is_file():
        pending.append(f"candidate-final {phase} final capture is missing")
        return False
    try:
        state = obj(load(state_path, str(state_path)), str(state_path))
        status = state.get("status")
        if status == "running":
            pending.append(f"candidate-final {phase} final capture is still running")
            return False
        if status != "pass":
            failures.append(f"candidate-final {phase} final capture is not a pass")
            return False
        if state.get("phase") != phase or state.get("attempt") != "final" or state.get("completed_lanes") != 6 or state.get("source_unchanged") is not True or state.get("outside_bundle_status_unchanged") is not True:
            failures.append(f"candidate-final {phase} final capture custody differs")
            return False
        index = state.get("index")
        if not isinstance(index, list) or len(index) != 6:
            failures.append(f"candidate-final {phase} final capture index differs")
            return False
        return True
    except (OSError, KeyError, TypeError, ValueError, VerificationError) as error:
        failures.append(str(error))
        return False


def check_final_candidate_epoch(
    precleanup: bool,
    receipts: dict[str, dict[str, Any]],
    final_source: dict[str, Any] | None,
    pending: list[str],
    failures: list[str],
) -> dict[str, Any] | None:
    """Replay the final candidate once both serialized phases are complete."""

    final_root = ROOT / "candidate-final"
    protocol_path = final_root / "protocol.json"
    if not protocol_path.is_file():
        pending.append("candidate-final protocol is missing")
        return None
    try:
        protocol = obj(load(protocol_path, str(protocol_path)), str(protocol_path))
        if protocol.get("schema") != "litchi-0457-odp-source-tail-candidate-v1" or protocol.get("change") != CHANGE:
            fail("candidate-final protocol schema/change differs")
        if protocol.get("capture_driver_sha256") != sha_file(final_root / "capture.py"):
            fail("candidate-final capture driver hash differs from its protocol")
    except (OSError, KeyError, TypeError, ValueError, VerificationError) as error:
        failures.append(str(error))
        return None
    if final_source is None:
        return None
    for name in FINAL_SOURCE_GATE_NAMES:
        receipt = receipts.get(name)
        if receipt is not None and receipt.get("status") == "pass" and receipt.get("source_after") != final_source:
            failures.append(f"checks/{name}.json: final source epoch differs from candidate-final build")
    if not all(final_capture_complete(phase, pending, failures) for phase in ("R1", "R2")):
        return None
    summary_path = final_root / "summary.json"
    comparison_path = ROOT / "comparison-final.json"
    if not summary_path.is_file():
        pending.append("candidate-final summary.json is missing")
        return None
    if not comparison_path.is_file():
        pending.append("comparison-final.json is missing")
        return None
    child_args = ["--attempt", "final"]
    if precleanup:
        child_args.extend(["--precleanup", "--repo-root", str(ROOT.parents[3])])
    try:
        child = delegate(final_root / "verify.py", child_args, "candidate-final verifier")
        derive_args = ["--root", str(final_root), "--protocol", str(protocol_path), "--attempt", "final"]
        if precleanup:
            derive_args.extend(["--precleanup", "--repo-root", str(ROOT.parents[3])])
        compare_derived(summary_path, final_root / "derive.py", derive_args, "candidate-final derive")
        compare_derived(
            comparison_path,
            ROOT / "comparison.py",
            ["--control", str(ROOT / "control/measurements.json"), "--candidate", str(summary_path)],
            "comparison-final derive",
            normalize_paths=True,
            input_paths={"control": "control/measurements.json", "candidate": "candidate-final/summary.json"},
        )
        return child
    except (OSError, KeyError, TypeError, ValueError, VerificationError) as error:
        failures.append(str(error))
        return None


def check_summary_file(path: Path, expected_schema: str, selector: str, protocol_path: Path, pending: list[str], failures: list[str]) -> dict[str, Any] | None:
    if not path.is_file():
        pending.append(f"missing derived summary: {path.relative_to(ROOT)}")
        return None
    summary = obj(load(path, str(path)), str(path))
    if summary.get("schema") != expected_schema or summary.get("change") != CHANGE:
        failures.append(f"{path.relative_to(ROOT)} schema/change differs")
    matrix = obj(summary.get("matrix"), f"{path}.matrix")
    expected = {"reports": 12, "retained_samples": 360, "samples_per_report": 30, "warmups_per_report": 3, "cpu": 2, "workers": 1, "phases": ["R1", "R2"], "selector": selector}
    for key, value in expected.items():
        if matrix.get(key) != value:
            failures.append(f"{path.relative_to(ROOT)} matrix.{key} differs")
    if summary.get("protocol_sha256") != sha_file(protocol_path):
        failures.append(f"{path.relative_to(ROOT)} protocol hash differs")
    rows = summary.get("rows")
    if not isinstance(rows, list) or len(rows) != 12:
        failures.append(f"{path.relative_to(ROOT)} does not contain twelve rows")
    return summary


def derived_json(script: Path, args: list[str], label: str) -> dict[str, Any]:
    result = subprocess.run([sys.executable, "-B", str(script), *args], cwd=ROOT, capture_output=True, text=True)
    if result.returncode != 0:
        fail(f"{label} failed: {result.stderr.strip() or result.stdout.strip()}")
    try:
        return obj(json.loads(result.stdout), label)
    except json.JSONDecodeError as error:
        fail(f"{label} did not return JSON: {error}")
    raise AssertionError("unreachable")


def compare_derived(path: Path, script: Path, args: list[str], label: str, *, normalize_paths: bool = False, input_paths: dict[str, str] | None = None) -> None:
    expected = obj(load(path, str(path)), str(path))
    actual = derived_json(script, args, label)
    if normalize_paths:
        # comparison.py records the absolute paths it was given.  Preserve
        # that provenance in the bound file while checking that a copied
        # bundle's inputs still resolve to its own relative artifacts.
        expected_inputs = obj(expected.get("inputs"), f"{label}.expected.inputs")
        actual_inputs = obj(actual.get("inputs"), f"{label}.actual.inputs")
        paths = input_paths or {"control": "control/measurements.json", "candidate": "candidate/summary.json"}
        for role, relative in paths.items():
            expected_role = obj(expected_inputs.get(role), f"{label}.expected.inputs.{role}")
            actual_role = obj(actual_inputs.get(role), f"{label}.actual.inputs.{role}")
            observed = Path(text(actual_role.get("path"), f"{label}.actual.inputs.{role}.path"))
            if not observed.is_absolute() or observed.resolve() != (ROOT / relative).resolve():
                fail(f"{label}: recomputed {role} path does not resolve to the retained bundle artifact")
            actual_role["path"] = expected_role.get("path")
    if actual != expected:
        failures = f"{label}: recomputed value differs from retained JSON"
        fail(failures)


def check_summaries(pending: list[str], failures: list[str], precleanup: bool) -> None:
    candidate = check_summary_file(ROOT / "candidate/summary.json", "litchi-0457-odp-source-tail-candidate-summary-v1", "odp_source_tail_append_lifecycle", ROOT / "candidate/protocol.json", pending, failures)
    control = check_summary_file(ROOT / "control/measurements.json", "litchi-0457-odp-control-summary-v1", "odp_existing_append_lifecycle", ROOT / "control/protocol-r1.json", pending, failures)
    if candidate is not None:
        candidate_args = ["--root", str(ROOT / "candidate"), "--protocol", str(ROOT / "candidate/protocol.json"), "--attempt", "formal"]
        # The historical candidate shares the Cargo target directory with the
        # final source epoch.  Its retained task-owned copies are checked by
        # check_build_bindings; do not ask the child derivation replay to
        # inspect a source executable that the final build may have replaced.
        try:
            compare_derived(ROOT / "candidate/summary.json", ROOT / "candidate/derive.py", candidate_args, "candidate derive")
        except (OSError, KeyError, TypeError, ValueError, VerificationError) as error:
            failures.append(str(error))
    if control is not None:
        control_args = ["--root", str(ROOT / "control"), "--protocol", str(ROOT / "control/protocol-r1.json"), "--attempt", "formal-r1"]
        try:
            compare_derived(ROOT / "control/measurements.json", ROOT / "control/derive.py", control_args, "control derive")
        except (OSError, KeyError, TypeError, ValueError, VerificationError) as error:
            failures.append(str(error))
    if candidate is not None and candidate.get("comparison_contract", {}).get("status") != "withheld":
        failures.append("candidate summary does not withhold ordinary comparison claims")
    comparison_path = ROOT / "comparison.json"
    if not comparison_path.is_file():
        pending.append("comparison.json is missing")
    else:
        comparison = obj(load(comparison_path, str(comparison_path)), "comparison.json")
        if comparison.get("schema") != "litchi-0457-odp-control-candidate-comparison-v1" or comparison.get("change") != CHANGE:
            failures.append("comparison schema/change differs")
        matrix = obj(comparison.get("matrix"), "comparison.matrix")
        if matrix.get("rows") != 12 or matrix.get("samples_per_row") != 30:
            failures.append("comparison matrix differs")
        if not isinstance(comparison.get("rows"), list) or len(comparison["rows"]) != 12:
            failures.append("comparison does not contain twelve rows")
        if candidate is not None and control is not None:
            try:
                compare_derived(comparison_path, ROOT / "comparison.py", ["--control", str(ROOT / "control/measurements.json"), "--candidate", str(ROOT / "candidate/summary.json")], "comparison derive", normalize_paths=True)
            except (OSError, KeyError, TypeError, ValueError, VerificationError) as error:
                failures.append(str(error))


PROFILE_REQUIRED_ARTIFACTS = {
    "catalog",
    "perf_data",
    "perf_record_stdout",
    "perf_record_stderr",
    "resource",
    "report",
    "top_symbols",
    "perf_script",
    "stacks_folded",
}
PROFILE_OPTIONAL_ARTIFACTS = {
    "oracle",
    "perf_help_stdout",
    "perf_help_stderr",
    "perf_script_stderr",
    "top_symbols_stderr",
}


def bundle_file(record: Any, label: str) -> Path:
    row = obj(record, label)
    path = safe_path(ROOT, row.get("path"), f"{label}.path")
    expected_sha = digest(row.get("sha256"), f"{label}.sha256")
    if sha_file(path) != expected_sha:
        fail(f"{label}: hash differs")
    if "bytes" in row and path.stat().st_size != integer(row.get("bytes"), f"{label}.bytes"):
        fail(f"{label}: byte count differs")
    return path


def profile_role_bindings(protocol: dict[str, Any], role_name: str, label: str) -> dict[str, Any]:
    role = obj(obj(protocol.get("roles"), f"{label}.roles").get(role_name), f"{label}.roles.{role_name}")
    for key in ("workload_protocol", "binary_binding", "build_receipt"):
        binding = obj(role.get(key), f"{label}.{key}")
        path = bundle_file(binding, f"{label}.{key}")
        if key == "workload_protocol" and path != ROOT / "candidate/protocol.json" and role_name == "candidate":
            fail(f"{label}: candidate workload protocol path differs")
    source = obj(role.get("source_manifest"), f"{label}.source_manifest")
    source_record(source, f"{label}.source_manifest")
    oracle = obj(role.get("oracle"), f"{label}.oracle")
    oracle_path = bundle_file(oracle, f"{label}.oracle") if "bytes" in oracle else safe_path(ROOT, oracle.get("verifier_path"), f"{label}.oracle.verifier_path")
    oracle_sha = digest(oracle.get("verifier_sha256"), f"{label}.oracle.verifier_sha256")
    if sha_file(oracle_path) != oracle_sha:
        fail(f"{label}: oracle verifier hash differs")
    return role


def profile_report(run_root: Path, receipt: dict[str, Any], role: dict[str, Any], label: str, revision: str) -> Path:
    artifacts = obj(receipt.get("artifacts"), f"{label}.artifacts")
    keys = set(artifacts)
    if not PROFILE_REQUIRED_ARTIFACTS.issubset(keys) or not keys.issubset(PROFILE_REQUIRED_ARTIFACTS | PROFILE_OPTIONAL_ARTIFACTS):
        fail(f"{label}: profile artifact set differs")
    for key, value in artifacts.items():
        artifact(run_root, value, f"{label}.artifacts.{key}")
    report_path = artifact(run_root, artifacts["report"], f"{label}.report")
    report = obj(load(report_path, str(report_path)), str(report_path))
    binary = obj(receipt.get("binary"), f"{label}.binary")
    identity = obj(report.get("binary_identity"), f"{label}.report.binary_identity")
    if identity.get("path") != binary.get("copy_path") or identity.get("binary_sha256") != binary.get("sha256") or identity.get("binary_bytes") != binary.get("bytes") or identity.get("profile") != "release" or identity.get("executable") is not True:
        fail(f"{label}: report binary identity differs")
    environment = obj(report.get("environment"), f"{label}.report.environment")
    if environment.get("git_revision") != revision or environment.get("rustflags") is not None:
        fail(f"{label}: report revision or rustflags differs")
    configuration = obj(report.get("configuration"), f"{label}.report.configuration")
    selector = text(role.get("selector"), f"{label}.selector")
    if configuration.get("samples_per_case") != 100 or configuration.get("warmup_iterations_per_case") != 3 or configuration.get("execution_workers") != [1] or configuration.get("cases") != [selector] or configuration.get("semantic_shapes") != ["large"]:
        fail(f"{label}: report dimensions differ")
    results = report.get("results")
    if not isinstance(results, list) or len(results) != 1:
        fail(f"{label}: expected one report result")
    result = obj(results[0], f"{label}.report.results[0]")
    if result.get("case") != selector or obj(result.get("corpus"), f"{label}.report.corpus").get("shape") != "large":
        fail(f"{label}: report selector or shape differs")
    elapsed = obj(result.get("elapsed_ns"), f"{label}.report.elapsed_ns")
    samples = elapsed.get("samples")
    if not isinstance(samples, list) or len(samples) != 100:
        fail(f"{label}: report elapsed vector is not 100 samples")
    operation = obj(result.get("operation_metrics"), f"{label}.report.operation_metrics")
    if operation.get("sample_count") != 100 or not isinstance(operation.get("sample_indices"), list) or sorted(operation["sample_indices"]) != list(range(100)):
        fail(f"{label}: report operation sample alignment differs")
    return report_path


def profile_oracle(role: dict[str, Any], report_path: Path, label: str, oracle_path: Path | None = None) -> None:
    oracle = obj(role.get("oracle"), f"{label}.oracle")
    verifier = oracle_path or safe_path(ROOT, oracle.get("verifier_path"), f"{label}.oracle.verifier_path")
    result = subprocess.run(
        [sys.executable, "-B", str(verifier), "--report", str(report_path), "--mode", "normal", "--shape", "large", "--samples", "100", "--warmups", "3"],
        cwd=ROOT,
        capture_output=True,
        text=True,
    )
    if result.returncode != 0 or result.stdout.strip() != "VALID":
        fail(f"{label}: independent profile oracle rejected report: {result.stderr.strip() or result.stdout.strip()}")


def profile_amendment(receipt_path: Path, report_path: Path, failures: list[str]) -> str:
    path = ROOT / "profiling/control-oracle-amendment.json"
    if not path.is_file():
        failures.append("profiling control oracle amendment is missing")
        return "failed"
    try:
        amendment = obj(load(path, str(path)), str(path))
        if amendment.get("schema") != "litchi-0457-profile-control-oracle-amendment-v1" or amendment.get("change") != CHANGE or amendment.get("status") != "validated":
            fail("profiling control oracle amendment schema/status differs")
        original = artifact(ROOT, amendment.get("original_profile"), "profiling amendment.original_profile")
        report = artifact(ROOT, amendment.get("report"), "profiling amendment.report")
        protocol = artifact(ROOT, amendment.get("applicable_control_protocol"), "profiling amendment.protocol")
        oracle = artifact(ROOT, amendment.get("oracle"), "profiling amendment.oracle")
        gate = artifact(ROOT, amendment.get("validation_gate"), "profiling amendment.validation_gate")
        artifact(ROOT, amendment.get("validation_log"), "profiling amendment.validation_log")
        if original != receipt_path or report != report_path:
            fail("profiling amendment paths do not bind the retained control receipt/report")
        receipt = obj(load(receipt_path, str(receipt_path)), "profiling control receipt")
        if sha_file(receipt_path) != digest(obj(amendment["original_profile"], "profiling amendment.original_profile").get("sha256"), "profiling amendment.original_profile.sha256"):
            fail("profiling amendment receipt hash differs")
        report_record = obj(amendment["report"], "profiling amendment.report")
        if report_record.get("sha256") != obj(receipt["artifacts"]["report"], "profiling control report artifact").get("sha256"):
            fail("profiling amendment report hash differs")
        if obj(amendment["applicable_control_protocol"], "profiling amendment.protocol").get("sha256") != sha_file(ROOT / "control/protocol-r1.json"):
            fail("profiling amendment protocol is not control/protocol-r1.json")
        if obj(amendment["oracle"], "profiling amendment.oracle").get("sha256") != sha_file(ROOT / "control/oracle/verify-report-r1.py"):
            fail("profiling amendment oracle hash differs")
        validation = obj(load(gate, str(gate)), "profiling amendment validation gate")
        if validation.get("status") != "pass" or validation.get("exit_code") != 0:
            fail("profiling amendment validation gate is not a pass")
        expected_argv = ["python3", "-B", "docs/performance/results/change-0457/control/oracle/verify-report-r1.py", "--report", "docs/performance/results/change-0457/profiling/r2/runs/control-large/report.json", "--mode", "normal", "--shape", "large", "--samples", "100", "--warmups", "3"]
        if validation.get("argv") != expected_argv:
            fail("profiling amendment validation argv differs")
        if amendment.get("samples") != 100 or amendment.get("warmups") != 3:
            fail("profiling amendment sample dimensions differ")
        result = subprocess.run([sys.executable, "-B", str(oracle), "--report", str(report), "--mode", "normal", "--shape", "large", "--samples", "100", "--warmups", "3"], cwd=ROOT, capture_output=True, text=True)
        if result.returncode != 0 or result.stdout.strip() != "VALID":
            fail("profiling amended control oracle rejected the retained report")
        return "accepted-amended-oracle"
    except (OSError, KeyError, TypeError, ValueError, VerificationError) as error:
        failures.append(str(error))
        return "failed"


def profile_receipt(protocol: dict[str, Any], protocol_path: Path, generation: str, role_name: str, expected_status: str, revision: str, precleanup: bool, pending: list[str], failures: list[str]) -> str:
    receipt_path = ROOT / "profiling" / generation / "runs" / f"{role_name}-large" / "profile-receipt.json"
    if not receipt_path.is_file():
        pending.append(f"profiling/{generation}/{role_name} profile receipt is missing")
        return "pending"
    label = f"profiling/{generation}/{role_name}"
    try:
        receipt = obj(load(receipt_path, str(receipt_path)), str(receipt_path))
        role = profile_role_bindings(protocol, role_name, label)
        if receipt.get("schema") != "litchi-0457-large-profile-receipt-v1" or receipt.get("change") != CHANGE or receipt.get("role") != role_name or receipt.get("label") != f"{role_name}-large-normal" or receipt.get("check_tag") != f"profile-{role_name}-large-{generation}":
            fail(f"{label}: receipt identity differs")
        if receipt.get("protocol_sha256") != sha_file(protocol_path) or receipt.get("runner_sha256") != sha_file(ROOT / f"profiling/capture-{generation}.py") or receipt.get("custody_wrapper_sha256") != sha_file(ROOT / "check.py"):
            fail(f"{label}: protocol/runner/custody hash differs")
        before = obj(receipt.get("source_before"), f"{label}.source_before")
        after = obj(receipt.get("source_after"), f"{label}.source_after")
        source_record(before, f"{label}.source_before")
        source_record(after, f"{label}.source_after")
        if before != after or receipt.get("source_unchanged") is not True:
            fail(f"{label}: source custody differs")
        provenance = obj(receipt.get("r1_provenance"), f"{label}.r1_provenance")
        if provenance.get("ambient_source_epoch") != before:
            fail(f"{label}: ambient source epoch differs")
        bound = obj(provenance.get("bound_source"), f"{label}.bound_source")
        expected_source = obj(role.get("source_manifest"), f"{label}.role.source_manifest")
        expected_bound = {"path": f"sources/{expected_source['sha256']}.json", "sha256": expected_source["sha256"], "files": expected_source["files"]}
        if bound != expected_bound:
            fail(f"{label}: bound source differs from role")
        source_record(bound, f"{label}.bound_source")
        for key in ("binary_binding", "build_receipt", "source_manifest", "workload_protocol"):
            actual_binding = obj(obj(receipt.get("bound_files"), f"{label}.bound_files").get(key), f"{label}.bound_files.{key}")
            expected_binding = obj(role.get(key), f"{label}.role.{key}")
            if any(actual_binding.get(field) != expected_binding.get(field) for field in ("path", "sha256")):
                fail(f"{label}: bound file {key} differs")
            if "files" in actual_binding and actual_binding.get("files") != expected_binding.get("files"):
                fail(f"{label}: bound file {key} differs")
        binary = obj(receipt.get("binary"), f"{label}.binary")
        role_binary = obj(role.get("binary"), f"{label}.role.binary")
        if {key: binary.get(key) for key in ("bytes", "sha256")} != {"bytes": role_binary.get("bytes"), "sha256": role_binary.get("sha256")} or binary.get("copy_path") != role_binary.get("path"):
            fail(f"{label}: binary binding differs")
        if "profile" in binary and binary.get("profile") != "release":
            fail(f"{label}: binary profile differs")
        if "executable" in binary and binary.get("executable") is not True:
            fail(f"{label}: binary executable flag differs")
        if receipt.get("allocator_environment") != {} or obj(receipt.get("counters"), f"{label}.counters").get("status") != "unavailable":
            fail(f"{label}: diagnostic allocator/counter policy differs")
        run_root = ROOT / "profiling" / generation
        report_path = profile_report(run_root, receipt, role, label, revision)
        commands = obj(receipt.get("commands"), f"{label}.commands")
        oracle_command = commands.get("oracle")
        if not isinstance(oracle_command, list) or "--report" not in oracle_command:
            fail(f"{label}: oracle command is malformed")
        report_arg = oracle_command[oracle_command.index("--report") + 1]
        if not Path(report_arg).as_posix().endswith(report_path.relative_to(ROOT).as_posix()):
            fail(f"{label}: oracle command report path differs")
        if generation == "r2" and (oracle_command.count("--samples") != 1 or oracle_command.count("--warmups") != 1):
            fail(f"{label}: amended oracle command must carry one sample/warmup override")
        status = receipt.get("status")
        if expected_status == "failed":
            if status != "failed" or receipt.get("record_exit_code") != 0 or not isinstance(receipt.get("oracle_exit_code"), int) or receipt.get("oracle_exit_code") == 0 or not receipt.get("failures"):
                fail(f"{label}: preserved failed profile identity differs")
            if generation == "r1" and ("--samples" in oracle_command or "--warmups" in oracle_command):
                fail(f"{label}: original failed oracle unexpectedly carries amended dimensions")
            return "authenticated-failure"
        if status == "pass":
            if receipt.get("record_exit_code") != 0 or receipt.get("oracle_exit_code") != 0 or receipt.get("failures"):
                fail(f"{label}: successful profile receipt has failure fields")
            profile_oracle(role, report_path, label)
            return "pass"
        if role_name == "control" and generation == "r2" and status == "failed":
            if receipt.get("record_exit_code") != 0 or receipt.get("oracle_exit_code") == 0:
                fail(f"{label}: control profile failure is not an oracle-only failure")
            return profile_amendment(receipt_path, report_path, failures)
        fail(f"{label}: unexpected profile status {status!r}")
    except (OSError, KeyError, TypeError, ValueError, VerificationError) as error:
        failures.append(str(error))
        return "failed"


def check_profiling(revision: str, precleanup: bool, pending: list[str], failures: list[str]) -> dict[str, str]:
    profile_root = ROOT / "profiling"
    if not profile_root.is_dir():
        return {}
    protocol_r2_path = profile_root / "protocol-r2.json"
    protocol_r1_path = profile_root / "protocol-r1.json"
    if not protocol_r2_path.is_file() or not protocol_r1_path.is_file():
        pending.append("profiling protocol amendment is incomplete")
        return {}
    try:
        protocol_r2 = obj(load(protocol_r2_path, str(protocol_r2_path)), "profiling protocol-r2")
        protocol_r1 = obj(load(protocol_r1_path, str(protocol_r1_path)), "profiling protocol-r1")
        for protocol, path in ((protocol_r1, protocol_r1_path), (protocol_r2, protocol_r2_path)):
            if protocol.get("schema") != "litchi-0457-odp-source-tail-profiling-v1" or protocol.get("change") != CHANGE or protocol.get("status") != "frozen-diagnostic-plan":
                failures.append(f"{path.relative_to(ROOT)} schema/status differs")
            scope = obj(protocol.get("scope"), f"{path}.scope")
            if scope.get("shape") != "large" or scope.get("mode") != "normal" or scope.get("cpu") != 2 or scope.get("workers") != 1 or scope.get("samples") != 100 or scope.get("warmups") != 3:
                failures.append(f"{path.relative_to(ROOT)} dimensions differ")
            custody = obj(protocol.get("custody"), f"{path}.custody")
            if custody.get("wrapper_path") != "check.py" or custody.get("wrapper_sha256") != sha_file(ROOT / "check.py") or custody.get("runner_sha256") != sha_file(ROOT / custody.get("runner_path")):
                failures.append(f"{path.relative_to(ROOT)} custody hashes differ")
        result = {
            "r1/candidate": profile_receipt(protocol_r1, protocol_r1_path, "r1", "candidate", "failed", revision, precleanup, pending, failures),
            "r2/candidate": profile_receipt(protocol_r2, protocol_r2_path, "r2", "candidate", "pass", revision, precleanup, pending, failures),
            "r2/control": profile_receipt(protocol_r2, protocol_r2_path, "r2", "control", "pass", revision, precleanup, pending, failures),
        }
        return result
    except (OSError, KeyError, TypeError, ValueError, VerificationError) as error:
        failures.append(str(error))
        return {}


def verify_sum_file(
    sum_path: Path, base: Path, excluded_prefix: str | None = None,
    excluded_files: frozenset[str] = frozenset(),
) -> None:
    rows: dict[str, str] = {}
    for line in sum_path.read_text(encoding="utf-8").splitlines():
        fields = line.split("  ", 1)
        if len(fields) != 2 or SHA256_RE.fullmatch(fields[0]) is None:
            fail(f"{sum_path}: malformed checksum line")
        name = fields[1]
        path = safe_path(base, name, f"{sum_path.name}.{name}")
        relative = path.relative_to(base).as_posix()
        if relative in rows or relative == sum_path.name:
            fail(f"{sum_path}: duplicate or self entry {relative}")
        if sha_file(path) != fields[0]:
            fail(f"{sum_path}: hash differs for {relative}")
        rows[relative] = fields[0]
    actual: set[str] = set()
    for path in base.rglob("*"):
        if path.is_symlink():
            fail(f"{base}: unsealed symlink {path.relative_to(base)}")
        if not path.is_file() or path == sum_path:
            continue
        relative = path.relative_to(base).as_posix()
        if relative in excluded_files:
            continue
        if excluded_prefix is None or not (relative == excluded_prefix or relative.startswith(excluded_prefix + "/")):
            actual.add(relative)
    if set(rows) != actual:
        fail(f"{sum_path}: checksum coverage differs")


def check_fuzz_static(pending: list[str], failures: list[str]) -> None:
    fuzz = ROOT / "fuzz"
    sums = fuzz / "SHA256SUMS"
    if not sums.is_file():
        pending.append("fuzz/SHA256SUMS is missing")
        return
    try:
        # This immutable seal records the inputs before fuzzing. The two
        # post-capture retention files are authenticated by the root seal and
        # checked against the original inventory by check_fuzz_post_run.
        verify_sum_file(
            sums, fuzz, "artifacts",
            frozenset({"retain-binaries.py", "binary-retention.json"}),
        )
    except VerificationError as error:
        failures.append(str(error))
    source_manifest = fuzz / "source-manifest.json"
    seed_manifest = fuzz / "seed-manifest.json"
    for path, kind in ((source_manifest, "source"), (seed_manifest, "seed")):
        if not path.is_file():
            failures.append(f"fuzz {kind}-manifest.json is missing")
            continue
        value = obj(load(path, str(path)), str(path))
        if value.get("change") != CHANGE:
            failures.append(f"fuzz {kind} manifest change differs")
        rows = value.get("source_rows" if kind == "source" else "seed_rows")
        if not isinstance(rows, list) or not rows:
            failures.append(f"fuzz {kind} manifest has no rows")
            continue
        seen: set[str] = set()
        for index, raw in enumerate(rows):
            row = obj(raw, f"fuzz {kind} row {index}")
            rel = text(row.get("path"), f"fuzz {kind} row {index}.path")
            if rel in seen:
                failures.append(f"fuzz {kind} manifest has duplicate {rel}")
            seen.add(rel)
            try:
                path_value = safe_path(fuzz, rel, f"fuzz {kind} row {index}.path")
                if path_value.stat().st_size != integer(row.get("bytes"), f"fuzz {kind} row {index}.bytes") or sha_file(path_value) != digest(row.get("sha256"), f"fuzz {kind} row {index}.sha256"):
                    failures.append(f"fuzz {kind} snapshot differs: {rel}")
            except VerificationError as error:
                failures.append(str(error))
    check_fuzz_post_run(pending, failures)


FUZZ_RETENTION_SCHEMA = "litchi-0457-fuzz-binary-retention-v1"


def decode_gzip(path: Path, expected_bytes: int, label: str) -> bytes:
    raw = path.read_bytes()
    if len(raw) > 128 * 1024 * 1024:
        fail(f"{label}: compressed artifact exceeds bound")
    stream = zlib.decompressobj(16 + zlib.MAX_WBITS)
    try:
        decoded = stream.decompress(raw, expected_bytes + 1)
        if len(decoded) <= expected_bytes:
            decoded += stream.flush()
    except zlib.error as error:
        fail(f"{label}: invalid gzip: {error}")
    if len(decoded) != expected_bytes or not stream.eof or stream.unused_data or stream.unconsumed_tail:
        fail(f"{label}: gzip decoded identity differs")
    return decoded


def check_fuzz_post_run(pending: list[str], failures: list[str]) -> None:
    fuzz = ROOT / "fuzz"
    post = fuzz / "artifacts/post-run"
    inventory_path = post / "inventory.json"
    if not inventory_path.is_file():
        pending.append("fuzz post-run inventory is missing")
        return
    inventory = obj(load(inventory_path, str(inventory_path)), "fuzz post-run inventory")
    if inventory.get("schema") != 1 or inventory.get("change") != CHANGE or inventory.get("runs") != 1000 or inventory.get("seed") != 457 or inventory.get("timeout_seconds") != 10:
        failures.append("fuzz post-run inventory dimensions differ")
    rows = inventory.get("files")
    if not isinstance(rows, list) or not rows:
        failures.append("fuzz post-run inventory has no files")
        return
    by_path: dict[str, dict[str, Any]] = {}
    for index, raw in enumerate(rows):
        row = obj(raw, f"fuzz post-run inventory.files[{index}]")
        rel = text(row.get("path"), f"fuzz post-run inventory.files[{index}].path")
        if rel in by_path:
            failures.append(f"fuzz post-run inventory duplicates {rel}")
        by_path[rel] = row
    # Large ASAN binaries are retained outside the post-run directory through
    # a separately sealed, lossless mapping. The capture inventory remains the
    # immutable source of the original bytes and digest.
    retention_path = fuzz / "binary-retention.json"
    retention: dict[str, dict[str, Any]] = {}
    if retention_path.is_file():
        value = obj(load(retention_path, str(retention_path)), "fuzz retention.json")
        if value.get("schema") != FUZZ_RETENTION_SCHEMA:
            failures.append("fuzz binary-retention manifest schema differs")
        try:
            inventory_ref = obj(value.get("original_inventory"), "fuzz binary-retention.original_inventory")
            inventory_ref_path = text(inventory_ref.get("path"), "fuzz binary-retention.original_inventory.path")
            if inventory_ref_path != "artifacts/post-run/inventory.json":
                failures.append("fuzz binary-retention inventory path differs")
            artifact(fuzz, inventory_ref, "fuzz binary-retention.original_inventory")
        except VerificationError as error:
            failures.append(str(error))
        try:
            driver_ref = obj(value.get("driver"), "fuzz binary-retention.driver")
            driver_path = text(driver_ref.get("path"), "fuzz binary-retention.driver.path")
            if driver_path != "retain-binaries.py":
                failures.append("fuzz binary-retention driver path differs")
            artifact(fuzz, driver_ref, "fuzz binary-retention.driver")
        except VerificationError as error:
            failures.append(str(error))
        records = value.get("records")
        if not isinstance(records, list) or not records:
            failures.append("fuzz binary-retention has no records")
            records = []
        retained_paths: set[str] = set()
        for index, raw in enumerate(records):
            row = obj(raw, f"fuzz retention row {index}")
            original_ref = obj(row.get("original"), f"fuzz retention row {index}.original")
            retained_ref = obj(row.get("retained"), f"fuzz retention row {index}.retained")
            original = text(original_ref.get("path"), f"fuzz retention row {index}.original.path")
            retained = text(retained_ref.get("path"), f"fuzz retention row {index}.retained.path")
            if original not in {"zip/parse_zip", "xml/scan_xml"}:
                failures.append(f"fuzz retention maps an unexpected binary: {original}")
            if original in retention or original not in by_path:
                failures.append(f"fuzz retention original path is duplicate or not inventoried: {original}")
                continue
            if retained in retained_paths:
                failures.append(f"fuzz retention retained path is duplicated: {retained}")
                continue
            retained_paths.add(retained)
            retention[original] = row
            try:
                retained_path = safe_path(post, retained, f"fuzz retention row {index}.retained_path")
                if retained_path == post / original:
                    failures.append(f"fuzz retention does not move {original}")
                if row.get("encoding") != "gzip":
                    failures.append(f"fuzz retention encoding differs: {original}")
                original_bytes = integer(original_ref.get("bytes"), f"fuzz retention row {index}.original.bytes")
                original_sha = digest(original_ref.get("sha256"), f"fuzz retention row {index}.original.sha256")
                retained_bytes = integer(retained_ref.get("bytes"), f"fuzz retention row {index}.retained.bytes")
                retained_sha = digest(retained_ref.get("sha256"), f"fuzz retention row {index}.retained.sha256")
                inventory_row = by_path[original]
                if original_bytes != integer(inventory_row.get("bytes"), f"fuzz inventory {original}.bytes") or original_sha != digest(inventory_row.get("sha256"), f"fuzz inventory {original}.sha256"):
                    failures.append(f"fuzz retention mapping disagrees with inventory: {original}")
                if retained_path.stat().st_size != retained_bytes or sha_file(retained_path) != retained_sha:
                    failures.append(f"fuzz retained compressed identity differs: {retained}")
                decoded = decode_gzip(retained_path, original_bytes, f"fuzz retained {retained}")
                if sha_file_bytes(decoded) != original_sha:
                    failures.append(f"fuzz retained decompressed identity differs: {original}")
            except VerificationError as error:
                failures.append(str(error))
    for rel, row in by_path.items():
        original = post / rel
        try:
            safe_path(post, rel, f"fuzz post-run inventory {rel}", must_exist=False)
        except VerificationError as error:
            failures.append(str(error))
            continue
        if original.is_symlink():
            failures.append(f"fuzz post-run inventory artifact is a symlink: {rel}")
            continue
        if original.is_file() and not original.is_symlink():
            try:
                artifact(post, row, f"fuzz post-run {rel}")
            except VerificationError as error:
                failures.append(str(error))
        elif rel in retention:
            # The retained gzip row was checked above.  The original path is
            # intentionally absent after lossless retention.
            mapped = retention[rel]
            original_ref = obj(mapped.get("original"), f"fuzz retention {rel}.original")
            if original_ref.get("sha256") != by_path[rel].get("sha256") or original_ref.get("bytes") != by_path[rel].get("bytes"):
                failures.append(f"fuzz retention mapping disagrees with inventory: {rel}")
        else:
            failures.append(f"fuzz post-run inventory artifact is missing: {rel}")
    actual_post: set[str] = set()
    for path in post.rglob("*"):
        if path.is_symlink():
            failures.append(f"fuzz post-run contains a symlink: {path.relative_to(post)}")
        elif path.is_file():
            actual_post.add(path.relative_to(post).as_posix())
    retained_paths = {obj(row.get("retained"), "fuzz retention retained").get("path") for row in retention.values()}
    expected_post = {"inventory.json"} | {
        rel for rel in by_path
        if rel not in retention or (post / rel).is_file()
    } | {str(path) for path in retained_paths if isinstance(path, str)}
    if actual_post != expected_post:
        failures.append("fuzz post-run contains uninventoried or missing retained files")
    for kind, binary in (("zip", "parse_zip"), ("xml", "scan_xml")):
        stage_dir = post / kind
        if not stage_dir.is_dir():
            failures.append(f"fuzz post-run {kind} directory is missing")
            continue
        corpus_inventory = stage_dir / "corpus-inventory.json"
        if not corpus_inventory.is_file():
            failures.append(f"fuzz post-run {kind} corpus inventory is missing")
        else:
            corpus = obj(load(corpus_inventory, str(corpus_inventory)), str(corpus_inventory))
            if corpus.get("kind") != kind or not isinstance(corpus.get("files"), list) or not corpus["files"]:
                failures.append(f"fuzz post-run {kind} corpus inventory is malformed")
            else:
                seen: set[str] = set()
                for index, raw in enumerate(corpus["files"]):
                    row = obj(raw, f"{kind} corpus row {index}")
                    name = text(row.get("path"), f"{kind} corpus row {index}.path")
                    if name in seen or Path(name).is_absolute() or "/" in name or ".." in Path(name).parts:
                        failures.append(f"{kind} corpus inventory has unsafe or duplicate name {name}")
                    seen.add(name)
                    integer(row.get("bytes"), f"{kind} corpus row {index}.bytes")
                    digest(row.get("sha256"), f"{kind} corpus row {index}.sha256")
        for phase in ("lock", "build", "smoke"):
            receipt_name = f"fuzz-0457-{kind}-{phase}.json"
            post_receipt = stage_dir / receipt_name
            root_receipt = ROOT / "checks" / receipt_name
            if not post_receipt.is_file() or not root_receipt.is_file():
                failures.append(f"fuzz {kind}/{phase} receipt is missing")
                continue
            if post_receipt.read_bytes() != root_receipt.read_bytes():
                failures.append(f"fuzz {kind}/{phase} post-run receipt differs from parent receipt")
            receipt = obj(load(root_receipt, str(root_receipt)), str(root_receipt))
            if receipt.get("status") != "pass" or receipt.get("exit_code") != 0 or receipt.get("source_unchanged") is not True:
                failures.append(f"fuzz {kind}/{phase} parent receipt is not a pass")
            try:
                root_log = obj(receipt.get("log"), f"{receipt_name}.log")
                root_log_path = artifact(ROOT, root_log, f"{receipt_name}.log")
                post_log_name = Path(text(root_log.get("path"), f"{receipt_name}.log.path")).name
                post_log = post / kind / post_log_name
                if not post_log.is_file() or post_log.read_bytes() != root_log_path.read_bytes():
                    failures.append(f"fuzz {kind}/{phase} post-run log differs from parent log")
            except VerificationError as error:
                failures.append(str(error))
        binary_path = stage_dir / binary
        if not binary_path.is_file() and f"{kind}/{binary}" not in retention:
            failures.append(f"fuzz {kind} binary is missing without retention mapping")


def sha_file_bytes(raw: bytes) -> str:
    return hashlib.sha256(raw).hexdigest()


def delegate(script: Path, args: list[str], label: str) -> dict[str, Any]:
    command = [sys.executable, "-B", str(script), *args]
    result = subprocess.run(command, cwd=ROOT, capture_output=True, text=True)
    if result.returncode != 0:
        fail(f"{label} failed: {result.stderr.strip() or result.stdout.strip()}")
    try:
        value = json.loads(result.stdout)
    except json.JSONDecodeError as error:
        fail(f"{label} did not return JSON: {error}")
    if not isinstance(value, dict) or value.get("status") != "pass":
        fail(f"{label} did not return a pass result")
    return value


def delegate_children(precleanup: bool, seal_ready: bool, pending: list[str], failures: list[str]) -> dict[str, Any]:
    results: dict[str, Any] = {}
    child_args = ["--precleanup"] if precleanup else []
    candidate_extra = ["--attempt", "formal"]
    for label, script, extra in (
        # The retained historical candidate is replayed after the outer
        # retention-copy check above.  It must not inspect the shared source
        # target, which belongs to a later final candidate epoch.
        ("candidate verifier", ROOT / "candidate/verify.py", candidate_extra),
        ("control verifier", ROOT / "control/verify.py", ["--attempt", "formal-r1", "--protocol", str(ROOT / "control/protocol-r1.json")]),
    ):
        if not script.is_file():
            pending.append(f"{label} is missing")
            continue
        try:
            role_args = extra if label == "candidate verifier" else child_args + extra
            results[label] = delegate(script, role_args, label)
        except VerificationError as error:
            failures.append(str(error))
    native = ROOT / "native/verify.py"
    if not native.is_file():
        pending.append("native verifier is missing")
    elif not seal_ready:
        pending.append("native verifier awaits the final root SHA256SUMS seal")
    else:
        try:
            results["native verifier"] = delegate(native, child_args, "native verifier")
        except VerificationError as error:
            failures.append(str(error))
    native_final = ROOT / "native-final/verify.py"
    if not native_final.is_file():
        pending.append("native-final verifier is missing")
    elif not seal_ready:
        pending.append("native-final verifier awaits the final root SHA256SUMS seal")
    else:
        try:
            results["native-final verifier"] = delegate(native_final, child_args, "native-final verifier")
        except VerificationError as error:
            failures.append(str(error))
    return results


def check_seal(pending: list[str], failures: list[str]) -> bool:
    seal = ROOT / "SHA256SUMS"
    if not seal.is_file():
        pending.append("change-0457/SHA256SUMS is missing")
        return False
    try:
        verify_sum_file(seal, ROOT)
    except VerificationError as error:
        failures.append(str(error))
        return False
    return True


def portable_replay(pending: list[str], failures: list[str]) -> None:
    try:
        with tempfile.TemporaryDirectory(prefix="litchi-0457-root-replay-") as directory:
            target = Path(directory) / ROOT.name
            shutil.copytree(ROOT, target, symlinks=True)
            result = subprocess.run([sys.executable, "-B", str(target / "verify.py")], cwd=target, capture_output=True, text=True)
            if result.returncode != 0:
                fail(f"portable copied verifier failed: {result.stderr.strip() or result.stdout.strip()}")
            value = json.loads(result.stdout)
            if not isinstance(value, dict) or value.get("status") != "pass" or value.get("portable") is not False:
                fail("portable copied verifier did not return an ordinary pass")
    except (OSError, json.JSONDecodeError, VerificationError) as error:
        failures.append(str(error))


def verify(precleanup: bool, portable: bool) -> dict[str, Any]:
    if precleanup and portable:
        fail("--precleanup and --portable are mutually exclusive")
    revision = parent_revision()
    receipts, pending, failures = audit_check_receipts(revision)
    selected = check_required_gates(receipts, pending, failures)
    check_final_gate_counts(receipts, failures)
    check_build_bindings(pending, failures, precleanup, revision)
    final_source = check_final_build_binding(receipts, pending, failures, revision)
    check_summaries(pending, failures, precleanup)
    profiling = check_profiling(revision, precleanup, pending, failures)
    final_candidate = check_final_candidate_epoch(precleanup, receipts, final_source, pending, failures)
    check_fuzz_static(pending, failures)
    seal_ready = check_seal(pending, failures)
    children = delegate_children(precleanup, seal_ready, pending, failures)
    if portable and not pending and not failures:
        portable_replay(pending, failures)
    result: dict[str, Any] = {
        "schema": SCHEMA,
        "change": CHANGE,
        "status": "fail" if failures else ("pending" if pending else "pass"),
        "precleanup": precleanup,
        "portable": portable,
        "parent_revision": revision,
        "selected_gates": selected,
        "profiling": profiling,
        "final_candidate": final_candidate,
        "failed_check_receipts": sorted(name for name, receipt in receipts.items() if receipt.get("status") == "failed"),
        "running_check_receipts": sorted(name for name, receipt in receipts.items() if receipt.get("status") == "running"),
        "child_results": children,
        "pending": sorted(set(pending)),
        "failures": sorted(set(failures)),
    }
    return result


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--precleanup", action="store_true")
    parser.add_argument("--portable", action="store_true")
    args = parser.parse_args(sys.argv[1:] if argv is None else argv)
    try:
        result = verify(args.precleanup, args.portable)
    except (OSError, KeyError, TypeError, ValueError, VerificationError) as error:
        result = {"schema": SCHEMA, "change": CHANGE, "status": "fail", "precleanup": args.precleanup, "portable": args.portable, "error": str(error)}
    print(json.dumps(result, ensure_ascii=False, sort_keys=True, indent=2))
    return 0 if result.get("status") == "pass" else (2 if result.get("status") == "pending" else 1)


if __name__ == "__main__":
    raise SystemExit(main())
