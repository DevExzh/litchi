#!/usr/bin/env python3
"""Run the complete untimed W8 File route matrix and emit a small receipt.

The route output is retained outside the checkout.  The committed receipt
should contain only the reduced per-route flags and the artifact identities;
the raw route JSON remains an external artifact.
"""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import subprocess
import sys

from analyze import semantic as analyze_semantic

EXPECTED_ROUTE_COUNT = 109
FULL_COMMIT = set("0123456789abcdef")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def artifact(path: Path) -> dict[str, object]:
    path = path.resolve()
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256(path)}


def hash_text(value: object, label: str) -> str:
    if not isinstance(value, str) or len(value) != 64:
        raise RuntimeError(f"{label} is not a SHA-256 hash")
    try:
        int(value, 16)
    except ValueError as error:
        raise RuntimeError(f"{label} is not a hexadecimal SHA-256 hash") from error
    return value


def run(binary: Path, route: str) -> tuple[dict, bytes]:
    command = [str(binary), "--mode", "validate", "--route", route]
    completed = subprocess.run(command, check=False, capture_output=True)
    if completed.returncode != 0:
        sys.stderr.buffer.write(completed.stderr)
        raise RuntimeError(f"{route} exited {completed.returncode}")
    try:
        value = json.loads(completed.stdout)
    except json.JSONDecodeError as error:
        raise RuntimeError(f"{route} returned invalid JSON: {error}") from error
    if not isinstance(value, dict) or value.get("route") != route:
        raise RuntimeError(f"{route} returned the wrong route identity")
    return value, completed.stdout


def validate(value: dict, route: str) -> dict[str, object]:
    validation = value.get("validation")
    if not isinstance(validation, dict):
        raise RuntimeError(f"{route} has no validation object")
    suffix = route.rsplit("/", 1)[-1]
    try:
        # Keep this runner aligned with the guarded qualification analyzer:
        # changed-route flags, repeated output identity and shared source
        # provenance must be checked before reducing the receipt.
        analyze_semantic(value, route, suffix)
    except SystemExit as error:
        raise RuntimeError(str(error)) from error
    for key in (
        "semantic_equal",
        "semantic_state_valid",
        "source_state_preserved",
        "source_state_exact",
        "exact_noop",
    ):
        if validation.get(key) is not True:
            raise RuntimeError(f"{route} validation.{key} is not true")
    for owner in ("leaf", "facade"):
        before = hash_text(
            validation.get(f"source_{owner}_before_sha256"),
            f"{route} source_{owner}_before_sha256",
        )
        after = hash_text(
            validation.get(f"source_{owner}_after_sha256"),
            f"{route} source_{owner}_after_sha256",
        )
        if before != after:
            raise RuntimeError(f"{route} source {owner} owner changed")
    outputs = (
        "source_leaf_output",
        "source_facade_output",
        "leaf_output",
        "facade_output",
        "leaf_repeat",
        "facade_repeat",
    )
    for key in outputs:
        output = validation.get(key)
        if not isinstance(output, dict) or not isinstance(output.get("bytes"), int) or output["bytes"] <= 0:
            raise RuntimeError(f"{route} {key} is not a nonempty output")
        hash_text(output.get("sha256"), f"{route} {key}.sha256")
        hash_text(output.get("source_sha256"), f"{route} {key}.source_sha256")
    if route.endswith("/open_get"):
        if validation.get("read_marker_equal") is not True:
            raise RuntimeError(f"{route} native/facade marker comparison failed")
        leaf = validation.get("leaf_read_marker_len")
        facade = validation.get("facade_read_marker_len")
        if not isinstance(leaf, int) or leaf < 0 or leaf != facade:
            raise RuntimeError(f"{route} native/facade marker lengths differ")
    return {
        "route": route,
        "semantic_equal": validation["semantic_equal"],
        "semantic_state_valid": validation["semantic_state_valid"],
        "source_state_preserved": validation["source_state_preserved"],
        "source_state_exact": validation["source_state_exact"],
        "exact_noop": validation["exact_noop"],
        "read_marker_equal": validation.get("read_marker_equal"),
        "leaf_read_marker_len": validation.get("leaf_read_marker_len"),
        "facade_read_marker_len": validation.get("facade_read_marker_len"),
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("binary", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--source", required=True, type=Path)
    parser.add_argument("--probe", required=True, type=Path)
    parser.add_argument("--lock", required=True, type=Path)
    parser.add_argument("--source-commit", required=True)
    parser.add_argument("--build-command", required=True)
    args = parser.parse_args()
    if (
        len(args.source_commit) != 40
        or set(args.source_commit) - FULL_COMMIT
    ):
        raise RuntimeError("--source-commit must be a full 40-character commit id")
    binary = args.binary.resolve()
    output = args.output.resolve()
    if output.exists() and any(output.iterdir()):
        raise RuntimeError(f"refusing to overwrite nonempty output {output}")
    output.mkdir(parents=True, exist_ok=True)
    listed = subprocess.run(
        [str(binary), "--mode", "list", "--route", "*"],
        check=False,
        capture_output=True,
    )
    if listed.returncode != 0:
        sys.stderr.buffer.write(listed.stderr)
        raise RuntimeError("route list failed")
    routes_value = json.loads(listed.stdout)
    routes = routes_value.get("routes") if isinstance(routes_value, dict) else None
    if not isinstance(routes, list) or not routes or not all(isinstance(route, str) for route in routes):
        raise RuntimeError("route list is malformed")
    if len(routes) != len(set(routes)):
        raise RuntimeError("route list contains duplicates")
    if len(routes) != EXPECTED_ROUTE_COUNT:
        raise RuntimeError(
            f"route list has {len(routes)} routes; expected {EXPECTED_ROUTE_COUNT}"
        )
    outcomes = []
    raw = output / "routes.jsonl"
    with raw.open("wb") as stream:
        for index, route in enumerate(routes):
            value, stdout = run(binary, route)
            outcome = validate(value, route)
            outcome["index"] = index
            outcomes.append(outcome)
            stream.write(stdout.rstrip(b"\n") + b"\n")
    receipt = {
        "schema": "litchi_file_untimed_validation.v1",
        "source_commit": args.source_commit,
        "build_command": args.build_command,
        "route_count": len(outcomes),
        "artifacts": {
            "source": artifact(args.source),
            "probe": artifact(args.probe),
            "binary": artifact(binary),
            "lock": artifact(args.lock),
            "routes": artifact(raw),
        },
        "outcomes": outcomes,
    }
    receipt_path = output / "receipt.json"
    receipt_path.write_text(json.dumps(receipt, indent=2) + "\n")
    print(json.dumps({"routes": len(outcomes), "receipt": str(receipt_path), "routes_artifact": str(raw)}))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:
        print(f"0362h untimed validation failed: {error}", file=sys.stderr)
        raise
