#!/usr/bin/env python3
"""Combine signed-native latency and external memory evidence.

Latency may be inspected independently, but this collector is the only
status that can say the signed-native qualification is complete.  It requires
the endpoint-bound ABBA receipts and the complete per-fixture memory summary.
"""
from __future__ import annotations

import argparse
import importlib.util
import json
from pathlib import Path
import sys
from typing import Any

HERE = Path(__file__).resolve().parent
SPEC = importlib.util.spec_from_file_location(
    "validate_signed_native", HERE / "validate_signed_native.py"
)
assert SPEC and SPEC.loader
VALIDATE = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(VALIDATE)


def canonical(value: object) -> bytes:
    return (
        json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
        + "\n"
    ).encode()


def load(path: Path, label: str) -> dict[str, Any]:
    path = Path(path)
    VALIDATE.reject_symlink_path(path, label)
    if not path.is_file():
        raise ValueError(f"{label} is not a regular file")
    return VALIDATE.load_json(path.resolve(), label)


def collect(
    latency_path: Path,
    memory_path: Path,
    before_manifest: Path,
    after_manifest: Path,
    output: Path,
) -> dict[str, Any]:
    before_manifest = Path(before_manifest)
    after_manifest = Path(after_manifest)
    VALIDATE.reject_symlink_path(before_manifest, "before endpoint manifest")
    VALIDATE.reject_symlink_path(after_manifest, "after endpoint manifest")
    after_manifest_resolved = after_manifest.resolve()
    latency = load(latency_path, "latency qualification")
    memory = load(memory_path, "memory qualification")
    VALIDATE.require_external(latency_path, "latency qualification")
    VALIDATE.require_external(memory_path, "memory qualification")
    if VALIDATE.canonical(memory) != memory_path.read_bytes():
        raise ValueError("memory qualification is not canonical JSON")
    before = VALIDATE.validate_endpoint_manifest(before_manifest, "before")
    after = VALIDATE.validate_endpoint_manifest(after_manifest, "after")
    pair = VALIDATE.validate_pair(before, after)
    VALIDATE.validate_abba(latency)
    fixture_manifest = VALIDATE.load_json(Path(after["fixtures"]["path"]), "after fixtures")
    VALIDATE.load_fixture_manifest(fixture_manifest, "after fixtures")
    fixture_names = {fixture["name"] for fixture in fixture_manifest["fixtures"]}
    memory_rows = memory.get("rows")
    memory_routes = tuple(sorted(VALIDATE.FINAL_ONLY))
    if not isinstance(memory_rows, list) or {row.get("route") for row in memory_rows if isinstance(row, dict)} != set(memory_routes):
        raise ValueError("memory summary does not contain every final signed timed route")
    if memory.get("endpoint", {}).get("manifest") != str(after_manifest_resolved):
        raise ValueError("memory summary endpoint manifest differs from final endpoint")
    VALIDATE.validate_memory_summary(
        memory,
        after,
        fixture_names,
        memory_routes,
        VALIDATE.artifact(after_manifest),
    )
    if memory.get("source") != latency.get("provenance", {}).get("pair"):
        raise ValueError("memory and latency endpoint pair provenance differs")
    output = Path(output)
    VALIDATE.reject_symlink_path(output, "final qualification output")
    output = output.resolve()
    repository = HERE.parents[3].resolve()
    if output == repository or repository in output.parents:
        raise ValueError("final qualification receipt must be external to the checkout")
    if output.exists():
        raise ValueError(f"refusing to overwrite final qualification receipt: {output}")
    result = {
        "schema": "litchi_signed_native_qualification_final.v1",
        "status": "complete",
        "source": pair,
        "before": {
            "manifest": VALIDATE.artifact(before_manifest),
            "source_commit": before["source_commit"],
            "binary": before["binary"],
        },
        "after": {
            "manifest": VALIDATE.artifact(after_manifest),
            "source_commit": after["source_commit"],
            "binary": after["binary"],
        },
        "latency": VALIDATE.artifact(latency_path),
        "memory": VALIDATE.artifact(memory_path),
        "memory_status": memory["status"],
        "routes": len(latency["routes"]),
        "memory_rows": len(memory["rows"]),
    }
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_bytes(canonical(result))
    return result


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--latency", required=True, type=Path)
    parser.add_argument("--memory", required=True, type=Path)
    parser.add_argument("--before-manifest", required=True, type=Path)
    parser.add_argument("--after-manifest", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()
    print(json.dumps(collect(args.latency, args.memory, args.before_manifest, args.after_manifest, args.output), indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:
        print(f"signed-native final collection failed: {error}", file=sys.stderr)
        raise
