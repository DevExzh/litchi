#!/usr/bin/env python3
"""Version-two adapter for the DrawingML sparse compile driver.

The actual source materialization and Cargo invocation remain the reviewed
driver implementation.  This adapter adds the v2 policy gate and rewrites
only the newly-created receipt's schema; it never changes a v1 receipt in
place and never accepts a caller-provided candidate identity as policy.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import subprocess
import sys
from typing import Any, Mapping, Sequence

# This adapter is an executable invoked by an absolute path from the source
# protocol.  Do not rely on a caller having preloaded ``tools`` on PYTHONPATH.
TOOLS_ROOT = Path(__file__).resolve().parents[1]
if str(TOOLS_ROOT) not in sys.path:
    sys.path.insert(0, str(TOOLS_ROOT))

from native_qualification import drawingml_compile_driver as legacy
from native_qualification import prepare_drawingml_raw_inputs_v2 as policy


SCHEMA = "litchi.drawingml.compile-preparation.v2"
PROTOCOL_VERSION = 2
FEATURES = legacy.FEATURES
REQUIRED_SOURCE_PATHS = legacy.REQUIRED_SOURCE_PATHS
PROBE_RELATIVE_PATH = legacy.PROBE_RELATIVE_PATH
ALLOCATOR_RELATIVE_PATH = legacy.ALLOCATOR_RELATIVE_PATH
DriverError = legacy.DriverError
_sha256 = legacy._sha256


def _historical_source_policy() -> dict[str, Any]:
    return {
        "ref": policy.HISTORICAL_SOURCE_REF,
        "path": policy.HISTORICAL_SOURCE_PATH,
        "blob_sha1": policy.HISTORICAL_SOURCE_BLOB_SHA1,
        "sha256": policy.HISTORICAL_SOURCE_SHA256,
        "bytes": policy.HISTORICAL_SOURCE_BYTES,
    }


def _receipt_policy(operation: str, role: str | None) -> dict[str, Any]:
    """Return the immutable v2 policy projection stored in compile receipts."""

    return {
        "schema": policy.SCHEMA,
        "operation": operation,
        "role": role,
        "candidate_binding": "checked-in-only",
        "historical_source": _historical_source_policy(),
        "role_locks": policy.ROLE_LOCKS,
        "candidate_manifests": list(policy.CANDIDATE_MANIFEST_PATHS),
        "toolchain_paths": list(policy.TOOLCHAIN_PATHS),
        "vendor_roots": list(policy.VENDOR_ROOTS),
        "probe_copy_paths": list(policy.PROBE_COPY_PATHS),
    }


def _toolchain_identity(root: Path) -> dict[str, dict[str, Any]]:
    identities: dict[str, dict[str, Any]] = {}
    for relative in policy.TOOLCHAIN_PATHS:
        digest, size = _sha256(root / relative)
        identities[relative] = {"path": relative, "sha256": digest, "bytes": size}
    return identities


def _driver_identity() -> dict[str, Any]:
    path = Path(__file__).resolve()
    payload = path.read_bytes()
    return {"path": str(path), "sha256": hashlib.sha256(payload).hexdigest(), "bytes": len(payload)}


def _policy_revision(revision: str, role: str | None = None) -> None:
    if len(revision) != 40 or revision != revision.lower() or any(char not in "0123456789abcdef" for char in revision):
        raise DriverError("source revision must be a full lowercase commit SHA")
    if role == "candidate":
        try:
            binding = policy._require_candidate_binding()
        except (OSError, ValueError) as error:
            raise DriverError(f"v2 candidate policy is not execution-ready: {error}") from error
        if revision != binding["revision"]:
            raise DriverError("candidate source revision differs from the checked-in v2 binding")
    elif revision not in policy.APPROVED_BASELINE_REVISIONS and revision != policy.HISTORICAL_SOURCE_REF:
        raise DriverError("source revision is outside the reviewed DrawingML v2 source fences")


def _validate_preparation(arguments: argparse.Namespace) -> dict[str, Any]:
    """Validate the complete v2 pair before invoking the legacy compiler."""

    if getattr(arguments, "role", None) not in {"baseline", "candidate"}:
        raise DriverError("v2 compile role must be baseline or candidate")

    required = (
        "preparation_receipt",
        "baseline_root",
        "candidate_root",
        "baseline_revision",
        "candidate_revision",
        "trusted_root",
    )
    if any(not isinstance(getattr(arguments, name, None), str) for name in required):
        raise DriverError("v2 compile requires its complete preparation receipt/source binding")
    try:
        receipt = policy.validate_preparation_receipt(
            Path(arguments.preparation_receipt),
            baseline_root=Path(arguments.baseline_root),
            candidate_root=Path(arguments.candidate_root),
            baseline_revision=arguments.baseline_revision,
            candidate_revision=arguments.candidate_revision,
            trusted_root=Path(arguments.trusted_root),
        )
    except (OSError, ValueError) as error:
        raise DriverError(f"v2 preparation receipt validation failed: {error}") from error
    expected_root = (
        Path(arguments.baseline_root).resolve()
        if arguments.role == "baseline"
        else Path(arguments.candidate_root).resolve()
    )
    actual_root = Path(arguments.source_root).resolve()
    expected_revision = (
        arguments.baseline_revision
        if arguments.role == "baseline"
        else arguments.candidate_revision
    )
    if actual_root != expected_root:
        raise DriverError("compile source root is not the source root bound by the v2 receipt")
    if arguments.source_revision != expected_revision:
        raise DriverError("compile source revision is not the source revision bound by the v2 receipt")
    identity = receipt.get(arguments.role)
    if not isinstance(identity, dict):
        raise DriverError(f"v2 preparation receipt has no {arguments.role} source identity")
    if identity.get("root") != str(actual_root) or identity.get("head") != expected_revision:
        raise DriverError(f"v2 {arguments.role} source identity differs from the compile request")
    return receipt


def _receipt_artifact(path: Path) -> dict[str, Any]:
    if path.is_symlink() or not path.is_file():
        raise DriverError(f"v2 preparation receipt is not a regular file: {path}")
    payload = path.read_bytes()
    try:
        value = json.loads(payload.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as error:
        raise DriverError(f"v2 preparation receipt is not valid JSON: {path}") from error
    if not isinstance(value, dict) or value.get("schema") != policy.SCHEMA or value.get("protocol_version") != 2:
        raise DriverError(f"v2 preparation receipt schema differs: {path}")
    return {
        "path": str(path.resolve()),
        "schema": policy.SCHEMA,
        "sha256": hashlib.sha256(payload).hexdigest(),
        "bytes": len(payload),
    }


def _rewrite_receipt(
    path: Path,
    operation: str,
    role: str | None,
    preparation_path: Path,
    preparation_receipt: Mapping[str, Any],
) -> None:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise DriverError(f"cannot read compile-driver receipt: {error}") from error
    if not isinstance(value, dict) or value.get("schema") != SCHEMA:
        raise DriverError("v2 compile-driver receipt has an unexpected schema")
    value["protocol_version"] = PROTOCOL_VERSION
    value["driver"] = _driver_identity()
    role_identity = preparation_receipt.get(role) if role else None
    if not isinstance(role_identity, dict):
        raise DriverError("v2 compile-driver receipt lacks its role source identity")
    value["source_identity"] = role_identity
    value["preparation"] = _receipt_artifact(preparation_path)
    value["policy"] = _receipt_policy(operation, role)
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def validate_compile_receipt(
    receipt_path: Path,
    *,
    preparation_receipt: Path,
    baseline_root: Path,
    candidate_root: Path,
    baseline_revision: str,
    candidate_revision: str,
    trusted_root: Path,
    role: str,
) -> dict[str, Any]:
    """Consume a v2 compile receipt without accepting a relabelled v1 file."""

    if role not in {"baseline", "candidate"}:
        raise DriverError("compile receipt role must be baseline or candidate")
    values = type(
        "PreparationArguments",
        (),
        {
            "preparation_receipt": str(preparation_receipt),
            "baseline_root": str(baseline_root),
            "candidate_root": str(candidate_root),
            "baseline_revision": baseline_revision,
            "candidate_revision": candidate_revision,
            "trusted_root": str(trusted_root),
            "role": role,
            "source_root": str(baseline_root if role == "baseline" else candidate_root),
            "source_revision": baseline_revision if role == "baseline" else candidate_revision,
        },
    )()
    prepared = _validate_preparation(values)
    try:
        value = json.loads(receipt_path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise DriverError(f"cannot read v2 compile receipt: {error}") from error
    if not isinstance(value, dict) or value.get("schema") != SCHEMA or value.get("protocol_version") != PROTOCOL_VERSION:
        raise DriverError("compile receipt is not a v2 receipt")
    if value.get("operation") != "compile" or value.get("role") != role:
        raise DriverError("compile receipt operation/role differs")
    if value.get("driver") != _driver_identity():
        raise DriverError("compile receipt driver identity differs")
    if value.get("source_identity") != prepared.get(role):
        raise DriverError("compile receipt source identity differs from its v2 preparation")
    source = value.get("source")
    source_identity = prepared.get(role)
    if (
        not isinstance(source, dict)
        or not isinstance(source_identity, dict)
        or source.get("root") != source_identity.get("root")
        or source.get("revision") != source_identity.get("head")
    ):
        raise DriverError("compile receipt source differs from its v2 preparation")
    preparation_identity = value.get("preparation")
    expected_preparation = _receipt_artifact(preparation_receipt)
    if preparation_identity != expected_preparation:
        raise DriverError("compile receipt preparation identity differs")
    policy_value = value.get("policy")
    if policy_value != _receipt_policy("compile", role):
        raise DriverError("compile receipt v2 policy projection differs")
    return value


def _prepare_source(arguments: argparse.Namespace) -> int:
    _policy_revision(arguments.revision)
    original_schema = legacy.SCHEMA
    legacy.SCHEMA = SCHEMA
    try:
        result = legacy._prepare_source(arguments)
    finally:
        legacy.SCHEMA = original_schema
    if result == 0:
        source_root = Path(arguments.source_root).resolve()
        sparse = subprocess.run(
            [
                "git",
                "-C",
                str(source_root),
                "sparse-checkout",
                "add",
                "--no-cone",
                "/rust-toolchain.toml",
            ],
            check=False,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )
        if sparse.returncode != 0:
            raise DriverError(
                f"cannot materialize v2 toolchain policy: {sparse.stderr.strip()}"
            )
    # Source preparation is a v2 receipt even though it is not yet a paired
    # execution receipt.  It still cannot bind a candidate identity.
    path = Path(arguments.output).resolve()
    value = json.loads(path.read_text(encoding="utf-8"))
    value["protocol_version"] = PROTOCOL_VERSION
    value["driver"] = _driver_identity()
    source_root = Path(arguments.source_root).resolve()
    value["toolchain_paths"] = list(policy.TOOLCHAIN_PATHS)
    value["toolchain"] = _toolchain_identity(source_root)
    value["policy"] = _receipt_policy("prepare", None)
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    return result


def _compile(arguments: argparse.Namespace) -> int:
    prepared = _validate_preparation(arguments)
    _policy_revision(arguments.source_revision, arguments.role)
    original_schema = legacy.SCHEMA
    legacy.SCHEMA = SCHEMA
    try:
        result = legacy._compile(arguments)
    finally:
        legacy.SCHEMA = original_schema
    _rewrite_receipt(
        Path(arguments.output).resolve(),
        "compile",
        arguments.role,
        Path(arguments.preparation_receipt),
        prepared,
    )
    return result


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)
    prepare = subparsers.add_parser("prepare", help="create a sparse reviewed source worktree")
    prepare.add_argument("--git-root", required=True)
    prepare.add_argument("--revision", required=True)
    prepare.add_argument("--source-root", required=True)
    prepare.add_argument("--output", required=True)
    prepare.set_defaults(function=_prepare_source)
    compile_parser = subparsers.add_parser("compile", help="compile the minimal v2 probe package")
    compile_parser.add_argument("--source-root", required=True)
    compile_parser.add_argument("--source-revision", required=True)
    compile_parser.add_argument("--probe-root", required=True)
    compile_parser.add_argument("--features", default="")
    compile_parser.add_argument("--role", choices=("baseline", "candidate"), required=True)
    compile_parser.add_argument("--preparation-receipt", required=True)
    compile_parser.add_argument("--baseline-root", required=True)
    compile_parser.add_argument("--candidate-root", required=True)
    compile_parser.add_argument("--baseline-revision", required=True)
    compile_parser.add_argument("--candidate-revision", required=True)
    compile_parser.add_argument("--trusted-root", required=True)
    compile_parser.add_argument("--work-dir", required=True)
    compile_parser.add_argument("--target-dir", required=True)
    compile_parser.add_argument("--output", required=True)
    compile_parser.add_argument("--cargo", default="cargo")
    compile_parser.set_defaults(function=_compile)
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    arguments = _parser().parse_args(argv)
    try:
        return int(arguments.function(arguments))
    except DriverError as error:
        print(f"drawingml v2 compile driver: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
