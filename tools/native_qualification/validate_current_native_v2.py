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
    POLICY_SCHEMA,
    PREPARATION_SCHEMA,
    QUALIFICATION_SCHEMA,
    PolicyError,
    PolicyIncomplete,
    PolicyMismatch,
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
    try:
        validate_routes(str(receipt.get("lane")), value.get("routes", []))
    except PolicyError as error:
        raise ValidationError(str(error)) from error


def validate_preparation(path: Path) -> dict[str, Any]:
    """Admit a ready preparation receipt and revalidate both source roots."""

    receipt, receipt_sha256 = _load(Path(path), "preparation receipt")
    if receipt.get("schema") != PREPARATION_SCHEMA:
        raise ValidationError("preparation receipt schema is not current-native v2")
    if receipt.get("policy_schema") != POLICY_SCHEMA or receipt.get("policy_sha256") != TRUSTED_POLICY.digest():
        raise ValidationError("preparation receipt is not bound to the trusted current-source policy")
    _validate_manifest(receipt)
    _validate_tooling(receipt)
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
        for field in ("route", "kind", "program", "output", "role"):
            if row.get(field) != getattr(expected, field):
                raise ValidationError(f"qualification result {index} {field} differs from the policy")
        if row.get("status") not in {"ok", "refused"}:
            raise ValidationError(f"qualification result {index} has an invalid status")
        evidence = row.get("evidence")
        if not isinstance(evidence, Mapping):
            raise ValidationError(f"qualification result {index} is missing source evidence")
        expected_bindings = {
            side: hashlib.sha256(_canonical(prepared["receipt"]["identities"][side])).hexdigest()
            for side in ("before", "after")
        }
        if evidence.get("source_bindings") != expected_bindings:
            raise ValidationError(f"qualification result {index} source bindings differ from preparation")
        output_artifact = evidence.get("output_artifact")
        if not isinstance(output_artifact, Mapping):
            raise ValidationError(f"qualification result {index} is missing output artifact identity")
        relative = output_artifact.get("relative_path")
        output_path = output_artifact.get("path")
        if relative != expected.output or not isinstance(output_path, str) or not Path(output_path).is_absolute():
            raise ValidationError(f"qualification result {index} output path differs from the policy")
        try:
            actual_output = _file_identity(Path(output_path), f"qualification result {index} output")
        except (OSError, ValidationError) as error:
            raise ValidationError(str(error)) from error
        if dict(output_artifact) != {**actual_output, "relative_path": relative}:
            raise ValidationError(f"qualification result {index} output artifact changed")
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
