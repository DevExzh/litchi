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
import sys
from typing import Any, Sequence

from native_qualification import drawingml_compile_driver as legacy
from native_qualification import prepare_drawingml_raw_inputs_v2 as policy


SCHEMA = "litchi.drawingml.compile-preparation.v2"
PROTOCOL_VERSION = 2
FEATURES = legacy.FEATURES
REQUIRED_SOURCE_PATHS = legacy.REQUIRED_SOURCE_PATHS
PROBE_RELATIVE_PATH = legacy.PROBE_RELATIVE_PATH
ALLOCATOR_RELATIVE_PATH = legacy.ALLOCATOR_RELATIVE_PATH
DriverError = legacy.DriverError


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


def _rewrite_receipt(path: Path, operation: str, role: str | None = None) -> None:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise DriverError(f"cannot read compile-driver receipt: {error}") from error
    if not isinstance(value, dict) or value.get("schema") != legacy.SCHEMA:
        raise DriverError("legacy compile-driver receipt has an unexpected schema")
    value["schema"] = SCHEMA
    value["protocol_version"] = PROTOCOL_VERSION
    value["driver"] = _driver_identity()
    value["policy"] = {
        "schema": policy.SCHEMA,
        "operation": operation,
        "role": role,
        "candidate_binding": "checked-in-only",
    }
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def _prepare_source(arguments: argparse.Namespace) -> int:
    _policy_revision(arguments.revision)
    result = legacy._prepare_source(arguments)
    _rewrite_receipt(Path(arguments.output).resolve(), "prepare")
    return result


def _compile(arguments: argparse.Namespace) -> int:
    _policy_revision(arguments.source_revision, arguments.role)
    result = legacy._compile(arguments)
    _rewrite_receipt(Path(arguments.output).resolve(), "compile", arguments.role)
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
