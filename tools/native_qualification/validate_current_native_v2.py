#!/usr/bin/env python3
"""Validate current-source v2 preparation and qualification receipts.

The validator is intentionally separate from the historical legacy validator.
It consumes the fixed policy module, rechecks endpoint identities at use time,
and rejects blocked/pending receipts before looking at route results.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import sys
from typing import Any, Mapping, Sequence

from native_qualification.current_native_policy_v2 import (
    ABBA_ORDER,
    EXECUTION_RECEIPT_RELATIVE,
    EXECUTION_SCHEMA,
    POLICY_SCHEMA,
    PREPARATION_SCHEMA,
    QUALIFICATION_SCHEMA,
    PolicyError,
    PolicyIncomplete,
    PolicyMismatch,
    RUNNER_KIND,
    TRUSTED_POLICY,
    validate_routes,
    validate_source,
)


class ValidationError(ValueError):
    """A v2 receipt or route stream is not admissible."""


def _canonical(value: Any) -> bytes:
    return (json.dumps(value, sort_keys=True, separators=(",", ":"), ensure_ascii=True) + "\n").encode("utf-8")


def _load(path: Path, label: str) -> tuple[dict[str, Any], str]:
    lexical = path.absolute()
    cursor = Path(lexical.anchor)
    for component in lexical.parts[1:]:
        cursor /= component
        if cursor.is_symlink():
            raise ValidationError(f"{label} contains a symlink component: {cursor}")
    if path.is_symlink() or not path.is_file():
        raise ValidationError(f"{label} is not a regular file: {path}")
    try:
        raw = path.read_bytes()
        value = json.loads(raw.decode("utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise ValidationError(f"cannot read {label}: {error}") from error
    if not isinstance(value, dict):
        raise ValidationError(f"{label} must contain a JSON object")
    if _canonical(value) != raw:
        raise ValidationError(f"{label} is not canonical JSON")
    return value, hashlib.sha256(raw).hexdigest()


def _same_observation(expected: Mapping[str, Any], actual: Mapping[str, Any], label: str) -> None:
    for field in ("source_root", "commit", "git_tree", "tree_sha256", "clean"):
        if expected.get(field) != actual.get(field):
            raise ValidationError(f"{label}.{field} changed since preparation")
    if expected.get("tracked_tree") != actual.get("tracked_tree"):
        raise ValidationError(f"{label}.tracked_tree changed since preparation")
    expected_artifacts = expected.get("artifacts")
    actual_artifacts = actual.get("artifacts")
    if not isinstance(expected_artifacts, Mapping) or not isinstance(actual_artifacts, Mapping):
        raise ValidationError(f"{label}.artifacts is malformed")
    for name, recorded in expected_artifacts.items():
        current = actual_artifacts.get(name)
        if name in {"probes", "source_manifests"}:
            if recorded != current:
                raise ValidationError(f"{label}.{name} changed since preparation")
            continue
        if not isinstance(recorded, Mapping) or not isinstance(current, Mapping):
            raise ValidationError(f"{label}.{name} artifact is malformed")
        if recorded != current:
            raise ValidationError(f"{label}.{name} changed since preparation")
    if expected.get("fixtures") != actual.get("fixtures"):
        raise ValidationError(f"{label}.fixtures changed since preparation")
    if expected.get("archive") != actual.get("archive"):
        raise ValidationError(f"{label}.archive changed since preparation")


def _file_identity(path: Path, label: str) -> dict[str, Any]:
    lexical = path.absolute()
    cursor = Path(lexical.anchor)
    for component in lexical.parts[1:]:
        cursor /= component
        if cursor.is_symlink():
            raise ValidationError(f"{label} contains a symlink component: {cursor}")
    if path.is_symlink() or not path.is_file():
        raise ValidationError(f"{label} is not a regular file: {path}")
    payload = path.read_bytes()
    return {"path": str(path.resolve()), "sha256": hashlib.sha256(payload).hexdigest(), "bytes": len(payload)}


def _validate_tooling(receipt: Mapping[str, Any]) -> None:
    tooling = receipt.get("tooling")
    if not isinstance(tooling, Mapping):
        raise ValidationError("preparation receipt tooling identities are missing")
    module_dir = Path(__file__).resolve().parent
    expected_paths = {
        "policy": module_dir / "current_native_policy_v2.py",
        "preparer": module_dir / "prepare_current_native_inputs_v2.py",
        "validator": module_dir / "validate_current_native_v2.py",
    }
    for name, expected_path in expected_paths.items():
        recorded = tooling.get(name)
        if not isinstance(recorded, Mapping) or dict(recorded) != _file_identity(expected_path, f"{name} tooling"):
            raise ValidationError(f"{name} tooling identity differs from the reviewed source")
    if receipt.get("preparer") != tooling.get("preparer"):
        raise ValidationError("preparation receipt preparer identity is not bound to tooling")


def _validate_manifest(receipt: Mapping[str, Any]) -> None:
    manifest = receipt.get("manifest")
    if not isinstance(manifest, Mapping):
        raise ValidationError("preparation receipt manifest identity is missing")
    path_value = manifest.get("path")
    if not isinstance(path_value, str) or not Path(path_value).is_absolute():
        raise ValidationError("preparation receipt manifest path is not absolute")
    path = Path(path_value)
    value, digest = _load(path, "qualification manifest")
    if digest != manifest.get("sha256") or len(_canonical(value)) != manifest.get("bytes"):
        raise ValidationError("qualification manifest changed since preparation")
    if value.get("schema") != PREPARATION_SCHEMA or value.get("policy_schema") != POLICY_SCHEMA:
        raise ValidationError("qualification manifest schema is not current-native v2")
    if value.get("lane") != receipt.get("lane"):
        raise ValidationError("qualification manifest lane differs from preparation receipt")
    if value.get("trusted_inputs") != TRUSTED_POLICY.manifest_binding():
        raise ValidationError("qualification manifest trusted source/lock/fixture binding differs from policy")
    try:
        validate_routes(str(receipt.get("lane")), value.get("routes", []))
    except PolicyError as error:
        raise ValidationError(str(error)) from error


def _validate_output_root(receipt: Mapping[str, Any], *, required: bool) -> Path | None:
    value = receipt.get("output_root")
    if value is None:
        if required:
            raise ValidationError("ready preparation is missing its disjoint output root")
        return None
    if not isinstance(value, str) or not Path(value).is_absolute():
        raise ValidationError("preparation output root must be an absolute path")
    root = Path(value)
    lexical = root.absolute()
    cursor = Path(lexical.anchor)
    for component in lexical.parts[1:]:
        cursor /= component
        if cursor.is_symlink():
            raise ValidationError(f"preparation output root contains a symlink component: {cursor}")
    if root.is_symlink() or not root.is_dir():
        raise ValidationError("preparation output root is not a regular directory")
    initial = receipt.get("output_root_initial")
    if not isinstance(initial, Mapping) or initial.get("path") != str(root.resolve()) or initial.get("entries") != []:
        raise ValidationError("preparation output root lacks a fresh-creation proof")
    if initial.get("sha256") != hashlib.sha256(_canonical([])).hexdigest():
        raise ValidationError("preparation output root creation proof is invalid")
    if receipt.get("execution_receipt_relative") != EXECUTION_RECEIPT_RELATIVE:
        raise ValidationError("preparation execution receipt path is not the fixed relative name")
    root_resolved = root.resolve()
    identities = receipt.get("identities")
    if isinstance(identities, Mapping):
        for side in ("before", "after"):
            identity = identities.get(side)
            if not isinstance(identity, Mapping):
                continue
            source_value = identity.get("source_root")
            if isinstance(source_value, str):
                source = Path(source_value).resolve()
                if root_resolved == source or root_resolved in source.parents or source in root_resolved.parents:
                    raise ValidationError("preparation output root overlaps a source root")
            archive = identity.get("archive")
            archive_value = archive.get("path") if isinstance(archive, Mapping) else None
            if isinstance(archive_value, str):
                archive_path = Path(archive_value).resolve()
                if root_resolved == archive_path or root_resolved in archive_path.parents or archive_path in root_resolved.parents:
                    raise ValidationError("preparation output root overlaps a source archive")
    return root_resolved


def _validate_output_artifact(evidence: Mapping[str, Any], expected: Any, output_root: Path, label: str) -> None:
    artifact = evidence.get("output_artifact")
    if not isinstance(artifact, Mapping):
        raise ValidationError(f"{label} is missing output artifact identity")
    relative = artifact.get("relative_path")
    if not isinstance(relative, str) or Path(relative).is_absolute() or ".." in Path(relative).parts:
        raise ValidationError(f"{label} output relative path is unsafe")
    if relative != expected.output:
        raise ValidationError(f"{label} output path differs from the policy")
    relative_path = Path(relative)
    expected_path = output_root / relative_path
    cursor = output_root
    for component in relative_path.parts:
        cursor /= component
        if cursor.is_symlink():
            raise ValidationError(f"{label} output path contains a symlink component: {cursor}")
    try:
        expected_path.absolute().relative_to(output_root.absolute())
    except ValueError as error:
        raise ValidationError(f"{label} output path escapes the prepared output root") from error
    expected_path = expected_path.absolute()
    recorded_path = artifact.get("path")
    if not isinstance(recorded_path, str) or Path(recorded_path).absolute() != expected_path:
        raise ValidationError(f"{label} output artifact is outside the prepared output root")
    try:
        actual_output = _file_identity(expected_path, label)
    except (OSError, ValidationError) as error:
        raise ValidationError(str(error)) from error
    if dict(artifact) != {**actual_output, "relative_path": relative}:
        raise ValidationError(f"{label} output artifact changed")


def _validate_abba_evidence(
    evidence: Mapping[str, Any],
    route: Any,
    output_root: Path,
    source_bindings: Mapping[str, str],
    label: str,
) -> None:
    legs = evidence.get("abba")
    if not isinstance(legs, list) or len(legs) != len(ABBA_ORDER):
        raise ValidationError(f"{label} ABBA evidence does not contain the fixed four legs")
    seen: set[str] = set()
    for index, (leg, record) in enumerate(zip(ABBA_ORDER, legs)):
        if not isinstance(record, Mapping):
            raise ValidationError(f"{label} ABBA leg {index} is malformed")
        if record.get("route") != route.route or record.get("role") != leg or record.get("index") != index:
            raise ValidationError(f"{label} ABBA leg {index} route/order differs from policy")
        if record.get("returncode") != 0 or record.get("source_binding") != source_bindings.get(leg):
            raise ValidationError(f"{label} ABBA leg {index} is not source-bound and successful")
        artifact = record.get("output_artifact")
        if not isinstance(artifact, Mapping):
            raise ValidationError(f"{label} ABBA leg {index} has no output artifact")
        relative = artifact.get("relative_path")
        if not isinstance(relative, str) or Path(relative).is_absolute() or ".." in Path(relative).parts or not relative.startswith(".abba/"):
            raise ValidationError(f"{label} ABBA leg {index} output path is unsafe")
        path = output_root / relative
        cursor = output_root
        for component in Path(relative).parts:
            cursor /= component
            if cursor.is_symlink():
                raise ValidationError(f"{label} ABBA leg {index} output contains a symlink component")
        try:
            path.absolute().relative_to(output_root.absolute())
        except ValueError as error:
            raise ValidationError(f"{label} ABBA leg {index} output escapes the prepared root") from error
        if not isinstance(artifact.get("path"), str) or Path(artifact["path"]).absolute() != path.absolute():
            raise ValidationError(f"{label} ABBA leg {index} output path is not bound")
        actual = _file_identity(path, f"{label} ABBA leg {index}")
        expected_identity = {**actual, "relative_path": relative}
        if dict(artifact) != expected_identity:
            raise ValidationError(f"{label} ABBA leg {index} output identity changed")
        if relative in seen:
            raise ValidationError(f"{label} ABBA output path is reused")
        seen.add(relative)


def _validate_execution_receipt(value: Mapping[str, Any], prepared: Mapping[str, Any]) -> None:
    output_root = prepared.get("output_root")
    if not isinstance(output_root, Path):
        raise ValidationError("qualification has no prepared output root for execution receipt")
    artifact = value.get("execution_receipt")
    if not isinstance(artifact, Mapping):
        raise ValidationError("qualification is missing the actual runner execution receipt")
    expected_path = output_root / EXECUTION_RECEIPT_RELATIVE
    if artifact.get("relative_path") != EXECUTION_RECEIPT_RELATIVE or artifact.get("path") != str(expected_path):
        raise ValidationError("execution receipt path is not the prepared output-root receipt")
    execution, digest = _load(expected_path, "runner execution receipt")
    actual_identity = {
        "path": str(expected_path),
        "sha256": digest,
        "bytes": expected_path.stat().st_size,
        "relative_path": EXECUTION_RECEIPT_RELATIVE,
    }
    if dict(artifact) != actual_identity:
        raise ValidationError("runner execution receipt identity changed")
    if execution.get("schema") != EXECUTION_SCHEMA or execution.get("status") != "completed":
        raise ValidationError("runner execution receipt is not a completed v2 execution")
    for field in ("policy_sha256", "preparation_receipt_sha256", "lane", "output_root", "routes", "source_artifacts", "tooling"):
        expected = value.get("policy_sha256") if field == "policy_sha256" else (
            prepared.get("receipt_sha256") if field == "preparation_receipt_sha256" else value.get(field)
        )
        if execution.get(field) != expected:
            raise ValidationError(f"runner execution receipt {field} is not bound to qualification")
    runner = execution.get("runner")
    expected_runner = _file_identity(Path(__file__).resolve().parents[1] / "native_qualification_runner.py", "runner")
    current_runner = _file_identity(Path(__file__).with_name("current_native_policy_v2.py"), "current-native v2 runner")
    if isinstance(runner, Mapping) and dict(runner) == current_runner:
        if execution.get("runner_kind") != RUNNER_KIND:
            raise ValidationError("current-native v2 runner kind is missing")
    elif not isinstance(runner, Mapping) or dict(runner) != expected_runner:
        raise ValidationError("runner execution receipt is not bound to a reviewed runner")
    if execution.get("abba_order") != list(ABBA_ORDER):
        raise ValidationError("runner execution receipt ABBA order differs from the fixed protocol")
    if execution.get("route_receipts") != value.get("results"):
        raise ValidationError("runner execution receipt route evidence differs from qualification")


def validate_preparation(path: Path) -> dict[str, Any]:
    """Admit a ready preparation receipt and revalidate both source roots."""

    receipt, receipt_sha256 = _load(Path(path), "preparation receipt")
    if receipt.get("schema") != PREPARATION_SCHEMA:
        raise ValidationError("preparation receipt schema is not current-native v2")
    if receipt.get("policy_schema") != POLICY_SCHEMA or receipt.get("policy_sha256") != TRUSTED_POLICY.digest():
        raise ValidationError("preparation receipt is not bound to the trusted current-source policy")
    _validate_manifest(receipt)
    _validate_tooling(receipt)
    output_root = _validate_output_root(receipt, required=TRUSTED_POLICY.status == "ready")
    TRUSTED_POLICY.require_execution_ready()
    if receipt.get("status") != "ready" or receipt.get("execution_admitted") is not True:
        raise ValidationError("preparation receipt is blocked and cannot be executed")
    lane = receipt.get("lane")
    if not isinstance(lane, str):
        raise ValidationError("preparation receipt lane is missing")
    try:
        validate_routes(lane, receipt.get("routes", []))
    except PolicyError as error:
        raise ValidationError(str(error)) from error
    identities = receipt.get("identities")
    if not isinstance(identities, Mapping):
        raise ValidationError("preparation receipt identities are missing")
    for side in ("before", "after"):
        recorded = identities.get(side)
        if not isinstance(recorded, Mapping):
            raise ValidationError(f"preparation receipt identity is missing: {side}")
        try:
            archive = recorded.get("archive")
            archive_path = Path(str(archive.get("path"))) if isinstance(archive, Mapping) and archive.get("path") else None
            current = validate_source(
                Path(str(recorded.get("source_root"))),
                "baseline" if side == "before" else "candidate",
                archive_path,
            )
        except PolicyError as error:
            raise ValidationError(str(error)) from error
        _same_observation(recorded, current, side)
    return {
        "receipt": receipt,
        "receipt_sha256": receipt_sha256,
        "lane": lane,
        "output_root": output_root,
    }


def validate_qualification(path: Path, preparation_path: Path) -> dict[str, Any]:
    """Validate a route-only qualification result against a ready receipt."""

    prepared = validate_preparation(preparation_path)
    value, value_sha256 = _load(Path(path), "qualification receipt")
    if value.get("schema") != QUALIFICATION_SCHEMA:
        raise ValidationError("qualification receipt schema is not current-native v2")
    if value.get("policy_schema") != POLICY_SCHEMA or value.get("policy_sha256") != TRUSTED_POLICY.digest():
        raise ValidationError("qualification receipt is not bound to the trusted policy")
    if value.get("preparation_receipt_sha256") != prepared["receipt_sha256"]:
        raise ValidationError("qualification receipt is not bound to its preparation receipt")
    lane = value.get("lane")
    if lane != prepared["lane"]:
        raise ValidationError("qualification lane differs from its preparation receipt")
    if value.get("abba_order") != list(ABBA_ORDER):
        raise ValidationError("qualification ABBA order differs from the fixed protocol")
    if prepared["output_root"] is None or value.get("output_root") != str(prepared["output_root"]):
        raise ValidationError("qualification output root is not the prepared disjoint root")
    if value.get("source_artifacts") != prepared["receipt"].get("identities"):
        raise ValidationError("qualification source/lock/archive artifacts differ from preparation")
    if value.get("tooling") != prepared["receipt"].get("tooling"):
        raise ValidationError("qualification tooling identities differ from preparation")
    try:
        validate_routes(lane, value.get("routes", []))
    except PolicyError as error:
        raise ValidationError(str(error)) from error
    rows = value.get("results")
    if not isinstance(rows, list) or len(rows) != len(TRUSTED_POLICY.routes_for(lane)):
        raise ValidationError("qualification result count differs from the fixed route contract")
    for index, row in enumerate(rows):
        if not isinstance(row, Mapping):
            raise ValidationError(f"qualification result {index} is not an object")
        expected = TRUSTED_POLICY.routes_for(lane)[index]
        for field in ("route", "kind", "program", "output", "role", "status"):
            if row.get(field) != getattr(expected, field):
                raise ValidationError(f"qualification result {index} {field} differs from the policy")
        evidence = row.get("evidence")
        if not isinstance(evidence, Mapping):
            raise ValidationError(f"qualification result {index} is missing source evidence")
        expected_bindings = {
            side: hashlib.sha256(_canonical(prepared["receipt"]["identities"][side])).hexdigest()
            for side in ("before", "after")
        }
        if evidence.get("source_bindings") != expected_bindings:
            raise ValidationError(f"qualification result {index} source bindings differ from preparation")
        _validate_output_artifact(evidence, expected, prepared["output_root"], f"qualification result {index}")
        _validate_abba_evidence(
            evidence,
            expected,
            prepared["output_root"],
            expected_bindings,
            f"qualification result {index}",
        )
    _validate_execution_receipt(value, prepared)
    return {"schema": QUALIFICATION_SCHEMA, "lane": lane, "policy_sha256": TRUSTED_POLICY.digest(), "sha256": value_sha256}


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("preparation", type=Path)
    parser.add_argument("qualification", type=Path, nargs="?")
    args = parser.parse_args(argv)
    try:
        if args.qualification is None:
            result = validate_preparation(args.preparation)
            print(f"validated current-native preparation {result['lane']}")
        else:
            result = validate_qualification(args.qualification, args.preparation)
            print(f"validated current-native qualification {result['lane']}")
    except (OSError, PolicyError, ValidationError) as error:
        print(f"current-native validation failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
