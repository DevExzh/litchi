#!/usr/bin/env python3
"""Check custody of change-0456 transient files before and after cleanup."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
from pathlib import Path
from typing import Any


TASK = Path("/tmp/litchi-goal-0456")
HEX64 = re.compile(r"^[0-9a-f]{64}$")


class CustodyError(ValueError):
    """A retained transient identity does not satisfy the custody contract."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise CustodyError(message)


def load(path: Path) -> Any:
    try:
        return json.loads(path.read_text())
    except (OSError, ValueError) as error:
        raise CustodyError(f"{path}: cannot read JSON: {error}") from error


def digest(path: Path) -> str:
    try:
        return hashlib.sha256(path.read_bytes()).hexdigest()
    except OSError as error:
        raise CustodyError(f"{path}: cannot hash file: {error}") from error


def record_identity(row: Any, label: str) -> tuple[int, str]:
    require(isinstance(row, dict), f"{label}: expected object")
    size = row.get("bytes")
    value = row.get("sha256")
    require(type(size) is int and size >= 0, f"{label}.bytes: expected non-negative integer")
    require(isinstance(value, str) and HEX64.fullmatch(value) is not None, f"{label}.sha256: expected SHA-256")
    return size, value


def absolute_task_path(path: Any, label: str) -> Path:
    require(isinstance(path, str) and path, f"{label}: expected path")
    value = Path(path)
    require(value.is_absolute(), f"{label}: transient path must be absolute")
    require(".." not in value.parts, f"{label}: path traversal")
    try:
        relative = value.relative_to(TASK)
    except ValueError as error:
        raise CustodyError(f"{label}: path escapes {TASK}") from error
    require(relative.parts and relative != Path("."), f"{label}: task directory is not a file path")
    return value


def relative_task_path(path: Any, label: str) -> Path:
    require(isinstance(path, str) and path, f"{label}: expected relative path")
    value = Path(path)
    require(not value.is_absolute(), f"{label}: inventory path must be relative")
    require(".." not in value.parts and value.parts and value != Path("."), f"{label}: path traversal")
    return value


def reject_symlink_chain(path: Path, label: str) -> None:
    current = path
    while True:
        require(not current.is_symlink(), f"{label}: symlink is not accepted")
        if current == TASK:
            return
        parent = current.parent
        require(parent != current, f"{label}: path is outside task")
        current = parent


def add_expected(expected: dict[Path, tuple[int, str, str]], path: Any, row: Any, label: str) -> None:
    target = absolute_task_path(path, label + ".path")
    size, value = record_identity(row, label)
    previous = expected.get(target)
    identity = (size, value, label)
    if previous is not None:
        require(previous[:2] == identity[:2], f"{label}: conflicting identity for {target}")
        return
    expected[target] = identity


def optional_json(path: Path) -> Any | None:
    return load(path) if path.is_file() and not path.is_symlink() else None


def collect_expected(root: Path) -> dict[Path, tuple[int, str, str]]:
    """Collect every transient path that the final evidence must account for."""
    expected: dict[Path, tuple[int, str, str]] = {}

    for path in sorted(root.glob("*-build.json")):
        value = load(path)
        require(isinstance(value, dict), f"{path.name}: expected build object")
        binaries = value.get("binaries")
        require(isinstance(binaries, dict), f"{path.name}.binaries: expected object")
        for name, row in binaries.items():
            require(isinstance(name, str) and name, f"{path.name}.binaries: invalid name")
            require(isinstance(row, dict), f"{path.name}.binaries.{name}: expected object")
            add_expected(expected, row.get("path"), row, f"{path.name}.binaries.{name}")

    probe = optional_json(root.parent / "probe-build.json")
    if probe is not None:
        require(isinstance(probe, dict), "probe-build.json: expected object")
        binary = probe.get("binary")
        add_expected(expected, binary.get("path") if isinstance(binary, dict) else None, binary, "probe-build.json.binary")

    for path in sorted(root.glob("*proof.json")):
        value = load(path)
        if not isinstance(value, dict) or not isinstance(value.get("rows"), list):
            continue
        for index, row in enumerate(value["rows"]):
            if not isinstance(row, dict):
                continue
            if isinstance(row.get("output"), dict):
                output = row["output"]
                add_expected(expected, output.get("path"), output, f"{path.name}.rows[{index}].output")
            if isinstance(row.get("report"), dict) and Path(row["report"].get("path", "")).is_absolute():
                report = row["report"]
                add_expected(expected, report.get("path"), report, f"{path.name}.rows[{index}].report")
            if isinstance(row.get("raw_profile"), dict):
                raw = row["raw_profile"]
                add_expected(expected, raw.get("path"), raw, f"{path.name}.rows[{index}].raw_profile")

    manifest_path = root.parent / "fuzz-inputs" / "manifest.json"
    manifest = load(manifest_path)
    require(isinstance(manifest, list), "fuzz-inputs/manifest.json: expected list")
    for index, row in enumerate(manifest):
        require(isinstance(row, dict), f"fuzz-inputs/manifest.json[{index}]: expected object")
        original = row.get("original_path")
        require(isinstance(original, str) and original, f"fuzz-inputs/manifest.json[{index}].original_path: missing")
        relative = Path(original)
        require(not relative.is_absolute() and ".." not in relative.parts, f"fuzz-inputs/manifest.json[{index}].original_path: path traversal")
        add_expected(expected, str(TASK / "fuzz" / relative), row, f"fuzz-inputs/manifest.json[{index}]")

    post_run_path = root.parent / "fuzz-inputs" / "post-run.json"
    post_run = load(post_run_path)
    require(isinstance(post_run, dict), "fuzz-inputs/post-run.json: expected object")
    artifacts = post_run.get("artifacts")
    require(isinstance(artifacts, list), "fuzz-inputs/post-run.json.artifacts: expected list")
    require(artifacts, "fuzz-inputs/post-run.json.artifacts: empty custody list")
    for index, row in enumerate(artifacts):
        add_expected(expected, row.get("path") if isinstance(row, dict) else None, row, f"fuzz-inputs/post-run.json.artifacts[{index}]")
    return expected


def check_precleanup(expected: dict[Path, tuple[int, str, str]]) -> None:
    require(TASK.is_dir() and not TASK.is_symlink(), f"{TASK}: transient task directory is unavailable")
    for path, (size, value, label) in expected.items():
        reject_symlink_chain(path, label)
        require(path.is_file(), f"{label}: transient file is unavailable: {path}")
        require(path.stat().st_size == size, f"{label}: byte count differs: {path}")
        require(digest(path) == value, f"{label}: SHA-256 differs: {path}")


def check_postcleanup(root: Path, expected: dict[Path, tuple[int, str, str]]) -> dict[str, Any]:
    require(not TASK.exists() and not TASK.is_symlink(), f"{TASK}: transient task directory remains")
    inventory_path = root.parent / "temporary-artifacts.json"
    inventory = load(inventory_path)
    require(isinstance(inventory, dict), "temporary-artifacts.json: expected object")
    require(inventory.get("task") == str(TASK), "temporary-artifacts.json.task: differs")
    rows = inventory.get("artifacts")
    require(isinstance(rows, list), "temporary-artifacts.json.artifacts: expected list")
    require(inventory.get("files") == len(rows), "temporary-artifacts.json.files: differs")
    total = 0
    retained: dict[str, tuple[int, str]] = {}
    for index, row in enumerate(rows):
        label = f"temporary-artifacts.json.artifacts[{index}]"
        require(isinstance(row, dict), f"{label}: expected object")
        relative = relative_task_path(row.get("path"), label + ".path")
        key = relative.as_posix()
        require(key not in retained, f"{label}: duplicate path")
        size, value = record_identity(row, label)
        retained[key] = (size, value)
        total += size
    require(inventory.get("bytes") == total, "temporary-artifacts.json.bytes: differs")

    for path, (size, value, label) in expected.items():
        relative = path.relative_to(TASK).as_posix()
        require(relative in retained, f"{label}: missing from cleanup inventory: {relative}")
        require(retained[relative] == (size, value), f"{label}: cleanup inventory identity differs: {relative}")
    return {"inventory": inventory_path.name, "files": len(rows), "bytes": total, "expected": len(expected)}


def check_transients(root: str | Path, precleanup: bool) -> dict[str, Any]:
    """Validate transient identities before cleanup or retained inventory after it."""
    formal = Path(root).resolve()
    require(formal.is_dir() and not formal.is_symlink(), f"formal evidence root is unavailable: {formal}")
    expected = collect_expected(formal)
    require(expected, "no transient custody rows were found")
    if precleanup:
        check_precleanup(expected)
        return {"status": "pass", "precleanup": True, "task": str(TASK), "expected": len(expected)}
    inventory = check_postcleanup(formal, expected)
    return {"status": "pass", "precleanup": False, "task": str(TASK), **inventory}


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("root", nargs="?", type=Path, default=Path(__file__).resolve().parent)
    parser.add_argument("--precleanup", action="store_true")
    arguments = parser.parse_args()
    print(json.dumps(check_transients(arguments.root, arguments.precleanup), sort_keys=True))
