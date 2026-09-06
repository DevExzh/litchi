#!/usr/bin/env python3
"""Construct the frozen 0437 ODP protocol from retained oracle inputs.

This is a source-only freeze driver.  It reads the copied candidate oracle,
completed pilot receipts, the explicit replay-driver set, and the current
0437 scratch directories, then writes one deterministic ``protocol.json``.
It never starts a workload, invokes Cargo, removes scratch data, or changes a
source checkout.  The final bundle should copy this file beside the protocol
before invoking it.
"""

from __future__ import annotations

import argparse
import copy
import hashlib
import json
from pathlib import Path
from typing import Any, Iterable


CHANGE = 437
ROLES = ("before-buffered", "after-buffered", "after-streaming")
MODES = ("normal", "allocator")
SHAPES = ("tiny", "medium", "large")
PHASE_ORDER = ("A1", "B1", "C1", "C2", "B2", "A2")
TMP_ROOT = Path("/tmp")
TMP_PREFIX = "litchi-goal-0437-"

# Keep this list explicit and in the same order as replay.py.  A freeze must
# fail if a named input is absent; silently inheriting a previous change's
# driver set would make the copied-bundle replay non-reproducible.
REPLAY_DRIVERS = (
    "replay.py",
    "check.py",
    "pilot.py",
    "save-binaries.py",
    "capture.py",
    "profile.py",
    "summary.py",
    "decision.py",
    "evidence-preflight.py",
    "final-evidence-preflight.py",
    "formal-suite.py",
    "before-hypothesis.py",
    "verify.py",
    "before-hypothesis-capture.py",
    "compare-strict.py",
    "compare-strict-candidate.py",
    "build-descriptors.py",
    "cleanup.py",
    "lifecycle.py",
    "portable-probes.py",
    "seal.py",
    "protocol.json",
    "protocol-draft.json",
    "planned-checks.json",
    "verify-report.py",
    "oracle-protocol.json",
    "verify-report-candidate.py",
    "pilot-candidate.py",
    "freeze.py",
)


class FreezeError(ValueError):
    pass


def fail(message: str) -> None:
    raise FreezeError(message)


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def load(path: Path, label: str) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"{label}: invalid JSON: {error}")
    if not isinstance(value, dict):
        fail(f"{label}: expected an object")
    return value


def write(path: Path, value: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def relative(bundle: Path, path: Path, label: str) -> str:
    try:
        value = path.resolve().relative_to(bundle.resolve()).as_posix()
    except ValueError:
        fail(f"{label}: path is outside the evidence bundle")
    if not value or value == "." or ".." in Path(value).parts:
        fail(f"{label}: unsafe relative path")
    return value


def bundle_file(bundle: Path, value: str, label: str) -> Path:
    path = Path(value)
    if path.is_absolute() or not value or ".." in path.parts:
        fail(f"{label}: must be a safe bundle-relative path")
    resolved = (bundle / path).resolve()
    if not resolved.is_relative_to(bundle.resolve()):
        fail(f"{label}: path escapes the bundle")
    if not resolved.is_file():
        fail(f"{label}: retained file is missing: {value}")
    return resolved


def safe_relative_list(values: Iterable[str], label: str) -> list[str]:
    result = list(values)
    if not result or len(set(result)) != len(result):
        fail(f"{label}: expected a nonempty list of unique paths")
    for value in result:
        path = Path(value)
        if path.is_absolute() or not value or ".." in path.parts:
            fail(f"{label}: unsafe bundle-relative path {value!r}")
    return result


def pilot_receipts() -> list[str]:
    return [
        f"pilots/{role}/initial/{mode}-{shape}-receipt.json"
        for role in ROLES
        for mode in MODES
        for shape in SHAPES
    ]


def preparatory_profile_receipts() -> list[str]:
    return [
        f"profiles/preparatory-before-buffered/initial/{kind}/receipt.json"
        for kind in ("stat", "record")
    ]


def formal_profile_receipts() -> list[str]:
    return [
        f"profiles/{role}/{kind}/receipt.json"
        for role in ROLES
        for kind in ("stat", "record")
    ]


def discover_cleanup_paths(explicit: list[str] | None) -> list[str]:
    if explicit:
        candidates = [Path(value) for value in explicit]
    else:
        candidates = sorted(
            (path for path in TMP_ROOT.glob(TMP_PREFIX + "*") if path.is_dir()),
            key=lambda path: path.name,
        )
    if len(candidates) != 12:
        fail(
            "cleanup_paths: expected exactly 12 current direct /tmp scratch directories, "
            f"found {len(candidates)}"
        )
    if len({str(path) for path in candidates}) != len(candidates):
        fail("cleanup_paths: duplicate directories")
    for path in candidates:
        if path.parent != TMP_ROOT or not path.name.startswith(TMP_PREFIX) or path.is_symlink():
            fail(f"cleanup_paths: unsafe scratch directory {path}")
    return [str(path) for path in candidates]


def infer_repo_root(bundle: Path, explicit: Path | None) -> Path:
    if explicit is not None:
        root = explicit.resolve()
    else:
        # This works for a copied bundle at docs/performance/results/change-0437.
        roots = [parent for parent in bundle.resolve().parents if (parent / "docs").is_dir()]
        if not roots:
            fail("repo_root: pass --repo-root when freezing outside the repository")
        root = roots[0]
    if not root.is_dir():
        fail(f"repo_root: directory is missing: {root}")
    return root


def preserved_paths(repo_root: Path, explicit: list[str] | None) -> list[str]:
    values = [Path(value) for value in explicit] if explicit else [repo_root / "target", repo_root / "tools/perf-baseline/target"]
    if not values or len({str(value) for value in values}) != len(values):
        fail("preserved_paths: expected unique directories")
    for path in values:
        if not path.is_dir() or path.is_symlink():
            fail(f"preserved_paths: directory is missing or unsafe: {path}")
    return [str(path) for path in values]


def validate_pilot_receipts(bundle: Path, names: list[str]) -> None:
    for name in names:
        path = bundle_file(bundle, name, "pilot receipt")
        row = load(path, name)
        if row.get("status") != "pass" or row.get("exit_code") != 0 or row.get("oracle_exit_code") != 0:
            fail(f"{name}: pilot is not a passing terminal receipt")
        role = row.get("role")
        if role not in ROLES:
            fail(f"{name}: pilot role is outside the frozen matrix")
        candidate = row.get("oracle_path") == "verify-report-candidate.py"
        if role == "before-buffered" and candidate:
            fail(f"{name}: before-buffered pilot must use the historical oracle")
        if role in {"after-buffered", "after-streaming"} and not candidate:
            fail(f"{name}: candidate-role pilot must use verify-report-candidate.py")


def validate_replay_drivers(bundle: Path, output_name: str) -> list[str]:
    names = safe_relative_list(REPLAY_DRIVERS, "replay_drivers")
    for name in names:
        if name == output_name:
            continue
        bundle_file(bundle, name, "replay driver")
    return names


def oracle_binding(bundle: Path, oracle_name: str, candidate_name: str) -> dict[str, Any]:
    oracle_path = bundle_file(bundle, oracle_name, "oracle protocol")
    candidate_path = bundle_file(bundle, candidate_name, "candidate oracle verifier")
    oracle = load(oracle_path, oracle_name)
    if oracle.get("change") != CHANGE:
        fail("oracle protocol: change must be 437")
    if oracle.get("samples") != 30 or oracle.get("warmups") != 3 or oracle.get("workers") != 1 or oracle.get("cpu") != 2:
        fail("oracle protocol: CPU/workers/sample dimensions differ from 0437")
    return {
        "path": oracle_name,
        "sha256": sha(oracle_path),
        "verifier_path": candidate_name,
        "verifier_sha256": sha(candidate_path),
        "roles": {role: role for role in ROLES},
    }


def construct(args: argparse.Namespace) -> dict[str, Any]:
    bundle = args.bundle.resolve()
    if not bundle.is_dir():
        fail(f"bundle: directory is missing: {bundle}")
    output_name = args.output.name
    if args.output.parent.resolve() != bundle:
        fail("output: must be directly inside the bundle")
    if output_name != "protocol.json":
        fail("output: final protocol must be named protocol.json")

    oracle_path = bundle_file(bundle, args.oracle_protocol, "oracle protocol")
    candidate_path = bundle_file(bundle, args.candidate_oracle, "candidate oracle verifier")
    source = load(oracle_path, args.oracle_protocol)
    binding = oracle_binding(bundle, args.oracle_protocol, args.candidate_oracle)

    pilot_names = pilot_receipts()
    validate_pilot_receipts(bundle, pilot_names)
    prep_names = preparatory_profile_receipts()
    formal_names = formal_profile_receipts()
    for name in prep_names + formal_names:
        safe_relative_list([name], "profile receipt")

    cleanup = discover_cleanup_paths(args.cleanup_path)
    repo_root = infer_repo_root(bundle, args.repo_root)
    preserved = preserved_paths(repo_root, args.preserved_path)
    goal = (args.goal_path or (repo_root / "docs" / "GOAL.md")).resolve()
    if not goal.is_file() or goal.is_symlink() or goal.name != "GOAL.md":
        fail(f"goal_path: pinned GOAL.md is missing or unsafe: {goal}")

    if sha(goal) != "bed4058bb76330daab8ce9d4bceff639ab3fbd7ea06634158bef41b133c4d1f1":
        fail("user GOAL.md differs from the session pin")
    replay = validate_replay_drivers(bundle, output_name)
    result = copy.deepcopy(source)
    result.update(
        {
            "status": "frozen",
            "protocol_schema": "litchi-0437-odp-protocol-v1",
            "classification": (
                "Matched ODP buffered control and candidate bounded streaming role; "
                "this evidence authorizes descriptive baseline comparison only."
            ),
            "oracle": binding,
            "oracle_roles": {role: role for role in ROLES},
            "preparatory_oracle_path": "verify-report.py",
            "preparatory_summary_module": "before-hypothesis",
            "preparatory_summary_path": "before-hypothesis.json",
            "summary_module": "summary",
            "summary_path": "summary.json",
            "retention_decision": {
                "driver": "decision.py",
                "driver_sha256": sha(bundle / "decision.py"),
                "path": "retention-decision.json",
                "large_peak_reduction_threshold_percent": 20.0,
                "scope": "Measured bounded publication enabler; disclose all latency and RSS flags independently.",
            },
            "candidate_build_descriptor": "after/build.json",
            "matrix": {
                "roles": list(ROLES),
                "phases": list(PHASE_ORDER),
                "formal_reports": 36,
                "retained_samples": 1080,
            },
            "pilot_contract": {
                "attempt": "initial",
                "samples": 3,
                "warmups": 1,
                "roles": list(ROLES),
                "receipts": pilot_names,
            },
            "pilots": {
                "required_receipts": pilot_names,
                "groups": {
                    "preparatory_before_buffered": 6,
                    "candidate_two_roles": 12,
                    "by_role": {role: 6 for role in ROLES},
                },
            },
            "preparatory_profiles": {
                "required_receipts": prep_names,
                "role": "before-buffered",
                "attempt": "initial",
                "module": "profile.py --preparatory",
            },
            "formal_profiles": {
                "required_receipts": formal_names,
                "roles": list(ROLES),
                "kinds": ["stat", "record"],
                "shape": "large",
            },
            "preparation": {
                "pilot_driver": "pilot.py",
                "candidate_pilot_driver": "pilot-candidate.py",
                "profile_driver": "profile.py",
                "oracle": "verify-report.py",
                "summary_module": "before-hypothesis",
                "summary_path": "before-hypothesis.json",
            },
            "same_identity_policy": copy.deepcopy(source.get("identity_scope", {})),
            "replay_drivers": replay,
            "cleanup_paths": cleanup,
            "cleanup_expected_paths": list(cleanup),
            "preserved_paths": preserved,
            "goal_path": str(goal),
            "goal_sha256": sha(goal),
            "precleanup_receipt": "checks/precleanup-portable.json",
            "cleanup_inventory": "checks/cleanup-inventory.json",
            "cleanup_receipt": "checks/task-cleanup.json",
            "probe_needles": {
                "cpu_affinity": "CPU affinity",
                "phase_timestamp": "formal lanes overlap or are out of capture order",
                "record_event": "perf-record sampling/call-graph binding is stale",
                "summary": "summary.json: retained summary differs from independent derivation",
                "missing_required": "file is missing",
                "semantic_digest": "semantic_sha256",
                "build_protocol": "protocol binding is stale",
                "compression": "original artifact binding differs",
            },
            "frozen_inputs": {
                "oracle_protocol": {"path": args.oracle_protocol, "sha256": binding["sha256"]},
                "candidate_oracle": {"path": args.candidate_oracle, "sha256": binding["verifier_sha256"]},
            },
        }
    )
    report_contract = copy.deepcopy(result.get("report_contract", {}))
    report_contract.update({"required_reports": 36, "retained_samples": 1080})
    result["report_contract"] = report_contract
    return result


def main(argv: list[str] | None = None) -> int:
    default_bundle = Path(__file__).resolve().parent
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--bundle", type=Path, default=default_bundle)
    parser.add_argument("--output", type=Path, default=None)
    parser.add_argument("--oracle-protocol", default="oracle-protocol.json")
    parser.add_argument("--candidate-oracle", default="verify-report-candidate.py")
    parser.add_argument("--repo-root", type=Path, default=None)
    parser.add_argument("--goal-path", type=Path, default=None)
    parser.add_argument("--preserved-path", action="append", default=None)
    parser.add_argument("--cleanup-path", action="append", default=None)
    args = parser.parse_args(argv)
    args.bundle = args.bundle.resolve()
    args.output = (args.output or (args.bundle / "protocol.json")).resolve()
    try:
        result = construct(args)
        if args.output.exists():
            fail(f"output already exists: {args.output}")
        write(args.output, result)
        print(args.output)
    except (OSError, TypeError, ValueError, FreezeError) as error:
        print(f"FREEZE INVALID: {error}")
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
