#!/usr/bin/env python3
"""Remove only the frozen 0437 ODP scratch paths after portable proof.

The list is deliberately read from the frozen protocol.  This draft refuses
to inherit a count or directory names from 0436: ``cleanup_paths`` must be
the exact ordered list and ``cleanup_expected_paths`` is the explicit
allowlist that authorizes it.  Every member must be a direct, non-symlink
child of ``/tmp`` whose name starts with ``litchi-goal-0437-``.
"""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
import shutil
from typing import Any


ROOT = Path(__file__).resolve().parent
CHANGE = 437
TMP_ROOT = Path("/tmp")
PREFIX = "litchi-goal-0437-"
PRECHECK = ROOT / "checks" / "precleanup-portable.json"
OUTPUT = ROOT / "checks" / "cleanup-inventory.json"


def digest(path: Path) -> str:
    value = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            value.update(block)
    return value.hexdigest()


def require(condition: bool, message: str) -> None:
    if not condition:
        raise RuntimeError(message)


def load_protocol() -> dict[str, Any]:
    path = ROOT / "protocol.json"
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise RuntimeError(f"invalid 0437 protocol: {error}") from error
    require(isinstance(value, dict) and value.get("change") == CHANGE, "protocol is not 0437")
    return value


def protocol_paths(value: dict[str, Any]) -> tuple[Path, ...]:
    paths = value.get("cleanup_paths")
    expected = value.get("cleanup_expected_paths")
    require(isinstance(paths, list) and paths, "protocol.cleanup_paths must be a nonempty list")
    require(
        isinstance(expected, list) and expected,
        "protocol.cleanup_expected_paths must be a nonempty explicit allowlist",
    )
    require(paths == expected, "cleanup_paths differs from the explicit cleanup allowlist")
    require(all(isinstance(path, str) and path for path in paths), "cleanup paths must be strings")
    require(len(set(paths)) == len(paths), "cleanup paths contain duplicates")

    result: list[Path] = []
    for raw in paths:
        path = Path(raw)
        require(path.is_absolute(), f"cleanup path is not absolute: {raw}")
        require(path.parent == TMP_ROOT, f"cleanup path is not a direct /tmp child: {raw}")
        require(path.name.startswith(PREFIX) and path.name != PREFIX, f"unsafe 0437 cleanup path: {raw}")
        require(".." not in path.parts, f"cleanup path contains traversal: {raw}")
        resolved_parent = path.parent.resolve()
        require(resolved_parent == TMP_ROOT, f"cleanup parent resolves outside /tmp: {raw}")
        result.append(path)
    return tuple(result)


def bundle_relative(value: Any, label: str) -> Path:
    require(isinstance(value, str) and value, f"{label} must be a nonempty bundle-relative path")
    path = Path(value)
    require(not path.is_absolute() and ".." not in path.parts, f"unsafe {label}: {value}")
    return ROOT / path


def preserved_paths(value: dict[str, Any]) -> tuple[Path, ...]:
    values = value.get("preserved_paths")
    require(isinstance(values, list) and values, "protocol.preserved_paths must be a nonempty list")
    require(all(isinstance(path, str) and path for path in values), "preserved paths must be strings")
    require(len(set(values)) == len(values), "preserved paths contain duplicates")
    result: list[Path] = []
    for raw in values:
        path = Path(raw)
        require(path.is_absolute() and ".." not in path.parts, f"unsafe preserved path: {raw}")
        result.append(path)
    return tuple(result)


def directory_identity(path: Path) -> tuple[int, int]:
    stat = path.stat()
    return stat.st_dev, stat.st_ino


def reject_links(path: Path) -> None:
    require(not path.is_symlink(), f"refusing to remove symlink: {path}")
    for child in path.rglob("*"):
        require(not child.is_symlink(), f"refusing scratch tree containing symlink: {child}")


def file_totals(path: Path) -> tuple[int, int]:
    files = [candidate for candidate in path.rglob("*") if candidate.is_file()]
    require(all(not candidate.is_symlink() for candidate in files), f"scratch tree contains a symlink: {path}")
    return len(files), sum(candidate.stat().st_size for candidate in files)


def main() -> int:
    value = load_protocol()
    paths = protocol_paths(value)
    preserved = preserved_paths(value)
    precheck = bundle_relative(value.get("precleanup_receipt", "checks/precleanup-portable.json"), "precleanup_receipt")
    output = bundle_relative(value.get("cleanup_inventory", "checks/cleanup-inventory.json"), "cleanup_inventory")
    goal = Path(value.get("goal_path", ""))
    goal_sha256 = value.get("goal_sha256")
    require(goal.is_absolute() and goal.name == "GOAL.md", "protocol.goal_path must be an absolute GOAL.md path")
    require(isinstance(goal_sha256, str) and len(goal_sha256) == 64, "protocol.goal_sha256 must be pinned")
    require(precheck.is_file(), f"{precheck}: portable precleanup gate is unavailable")
    precheck_value = json.loads(precheck.read_text(encoding="utf-8"))
    require(
        isinstance(precheck_value, dict)
        and precheck_value.get("status") == "pass"
        and precheck_value.get("exit_code") == 0,
        f"{precheck}: portable precleanup gate is not passing",
    )
    require(not output.exists(), f"{output}: cleanup inventory already exists")
    require(goal.is_file() and not goal.is_symlink(), f"pinned goal is unavailable: {goal}")
    goal_before = digest(goal)
    require(goal_before == goal_sha256, "docs/GOAL.md digest differs from the pinned goal")

    for path in preserved:
        require(path.is_dir() and not path.is_symlink(), f"preserved target is unavailable: {path}")
    before_targets = {str(path): list(directory_identity(path)) for path in preserved}

    removed: list[dict[str, Any]] = []
    for path in paths:
        require(path.is_dir() and not path.is_symlink(), f"cleanup directory is unavailable: {path}")
        reject_links(path)
        files, bytes_count = file_totals(path)
        removed.append(
            {"path": str(path), "regular_files": files, "regular_file_bytes": bytes_count}
        )

    for path in paths:
        shutil.rmtree(path)
        require(not path.exists(), f"cleanup directory remains after removal: {path}")

    after_targets = {str(path): list(directory_identity(path)) for path in preserved}
    require(after_targets == before_targets, "shared target directory identity changed")
    require(digest(goal) == goal_before == goal_sha256, "docs/GOAL.md changed during cleanup")

    record = {
        "change": CHANGE,
        "status": "pass",
        "protocol": "protocol.json",
        "precleanup_receipt": str(precheck.relative_to(ROOT)),
        "cleanup_paths": [str(path) for path in paths],
        "cleanup_expected_paths": [str(path) for path in paths],
        "removed": removed,
        "preserved_target_directory_identity": before_targets,
        "user_goal_path": str(goal),
        "user_goal_sha256": goal_sha256,
    }
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(record, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(
        json.dumps(
            {
                "status": "pass",
                "removed_directories": len(removed),
                "regular_file_bytes": sum(row["regular_file_bytes"] for row in removed),
            },
            sort_keys=True,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
