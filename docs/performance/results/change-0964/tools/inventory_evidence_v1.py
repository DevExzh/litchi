#!/usr/bin/env python3
"""Inventory final 0964 evidence without entering reproducible build trees.

The inventory is intentionally a closure-time operation.  It refuses the
default ``live`` state so a caller must explicitly attest that captures and
reports are terminal before hashing them.  It never runs a workload, build,
or report generator.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import stat
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable

DEFAULT_ROOT = Path("/home/zhuhe/litchi-performance-0964")
FIXED_EXCLUDED_DIRECTORY_NAMES = {"source-baseline", "source-candidate", "tmp", "__pycache__"}
COMMAND_KEYS = {
    "argv",
    "build_command",
    "capture_command",
    "command",
    "regenerate_command",
    "regeneration_command",
    "reducer_command",
}
DEFAULT_MAX_FILES = 250_000
DEFAULT_MAX_TOTAL_BYTES = 1 << 50


def fail(message: str) -> None:
    raise RuntimeError(message)


def digest_bytes(value: bytes) -> str:
    digest = hashlib.sha256()
    digest.update(value)
    return digest.hexdigest()


def digest_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def excluded_directory(name: str) -> bool:
    return name in FIXED_EXCLUDED_DIRECTORY_NAMES or name.startswith("target-")


def descriptor(path: Path, root: Path) -> dict[str, Any]:
    info = path.lstat()
    mode = stat.S_IMODE(info.st_mode)
    relative = path.relative_to(root).as_posix()
    if stat.S_ISREG(info.st_mode):
        return {
            "path": relative,
            "kind": "file",
            "bytes": info.st_size,
            "mode": mode,
            "sha256": digest_file(path),
        }
    if stat.S_ISLNK(info.st_mode):
        target = os.readlink(path)
        encoded = os.fsencode(target)
        return {
            "path": relative,
            "kind": "symlink",
            "bytes": len(encoded),
            "mode": mode,
            "sha256": digest_bytes(encoded),
            "target": target,
        }
    fail(f"unsupported filesystem object in evidence root: {path}")


def walk_evidence(root: Path, output: Path, max_files: int, max_total_bytes: int) -> list[dict[str, Any]]:
    if not root.is_dir() or root.is_symlink():
        fail(f"evidence root missing or symlinked: {root}")
    rows: list[dict[str, Any]] = []
    total_bytes = 0
    for directory, directory_names, file_names in os.walk(root, topdown=True, followlinks=False):
        directory_path = Path(directory)
        directory_names[:] = sorted(name for name in directory_names if not excluded_directory(name))
        for name in sorted(file_names):
            path = directory_path / name
            if path.resolve() == output.resolve():
                continue
            row = descriptor(path, root)
            rows.append(row)
            total_bytes += row["bytes"]
            if len(rows) > max_files:
                fail(f"evidence file count exceeds bound {max_files}")
            if total_bytes > max_total_bytes:
                fail(f"evidence logical bytes exceed bound {max_total_bytes}")
    rows.sort(key=lambda row: row["path"])
    return rows


def command_value(value: Any) -> str | list[str] | None:
    if isinstance(value, str) and value.strip():
        return value
    if isinstance(value, list) and value and all(isinstance(item, str) for item in value):
        return value
    return None


def find_command_references(value: Any, source: str, prefix: str = "") -> Iterable[dict[str, Any]]:
    if isinstance(value, dict):
        for key, child in value.items():
            field = f"{prefix}.{key}" if prefix else key
            if key.lower() in COMMAND_KEYS:
                command = command_value(child)
                if command is not None:
                    kind = "regeneration" if "regenerat" in key.lower() else "recorded_command"
                    yield {"source": source, "field": field, "kind": kind, "command": command}
            yield from find_command_references(child, source, field)
    elif isinstance(value, list):
        for index, child in enumerate(value):
            yield from find_command_references(child, source, f"{prefix}[{index}]")


def command_references(root: Path, rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    references: list[dict[str, Any]] = []
    for row in rows:
        if row.get("kind") != "file" or not row["path"].endswith(".json"):
            continue
        path = root / row["path"]
        try:
            value = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, UnicodeError, json.JSONDecodeError):
            continue
        references.extend(find_command_references(value, row["path"]))
    unique: dict[tuple[str, str, str], dict[str, Any]] = {}
    for reference in references:
        command = reference["command"]
        marker = json.dumps(command, sort_keys=True, separators=(",", ":"))
        unique[(reference["source"], reference["field"], marker)] = reference
    return [unique[key] for key in sorted(unique)]


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=DEFAULT_ROOT)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--capture-state", choices=("live", "stable"), default="live")
    parser.add_argument("--max-files", type=int, default=DEFAULT_MAX_FILES)
    parser.add_argument("--max-total-bytes", type=int, default=DEFAULT_MAX_TOTAL_BYTES)
    args = parser.parse_args()

    if args.capture_state != "stable":
        fail("refusing to inventory live capture files; pass --capture-state stable at closure")
    if args.max_files <= 0 or args.max_total_bytes <= 0:
        fail("inventory bounds must be positive")
    root = args.root.expanduser().resolve()
    output = args.output.expanduser().resolve()
    if output.exists() or output.is_symlink():
        fail(f"refusing to overwrite inventory: {output}")
    if output.parent == root or root in output.parents:
        output.parent.mkdir(parents=True, exist_ok=True)

    rows = walk_evidence(root, output, args.max_files, args.max_total_bytes)
    references = command_references(root, rows)
    result = {
        "schema": "litchi.performance.0964.evidence-inventory.v1",
        "status": "inventory_ready",
        "created_utc": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "root": str(root),
        "capture_state_attestation": "stable",
        "scan_policy": {
            "follow_symlinks": False,
            "excluded_directory_names": sorted(FIXED_EXCLUDED_DIRECTORY_NAMES),
            "excluded_directory_prefixes": ["target-"],
            "output_excluded": str(output),
            "max_files": args.max_files,
            "max_total_bytes": args.max_total_bytes,
            "descriptor": "regular files carry logical bytes, permission mode, and SHA-256; symlinks carry link-byte descriptors",
        },
        "files": rows,
        "file_count": len(rows),
        "total_bytes": sum(row["bytes"] for row in rows),
        "regeneration_command_references": references,
        "regeneration_reference_policy": {
            "source": "included JSON receipts/protocol artifacts",
            "recognized_keys": sorted(COMMAND_KEYS),
            "note": "References are recorded commands from custody artifacts; this helper does not execute them.",
        },
    }
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("x", encoding="utf-8") as stream:
        json.dump(result, stream, indent=2, sort_keys=True)
        stream.write("\n")
    print(json.dumps({"status": result["status"], "files": len(rows), "bytes": result["total_bytes"], "output": str(output)}, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (OSError, RuntimeError, ValueError) as exc:
        print(f"inventory_evidence_v1: FAIL: {exc}")
        raise SystemExit(2)
