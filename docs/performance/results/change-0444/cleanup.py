#!/usr/bin/env python3
"""Remove only explicitly bound 0444 temporary directories after proof."""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import shutil
from typing import Any

ROOT = Path(__file__).resolve().parent
TMP = Path("/tmp")
PREFIX = "litchi-goal-0444-"


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def load(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def require(condition: bool, message: str) -> None:
    if not condition:
        raise RuntimeError(message)


def bundle_path(value: Any, label: str) -> Path:
    require(isinstance(value, str) and value, f"{label} must be a nonempty bundle-relative path")
    path = Path(value)
    require(not path.is_absolute() and ".." not in path.parts, f"unsafe {label}")
    resolved = (ROOT / path).resolve()
    require(resolved.is_relative_to(ROOT.resolve()), f"{label} escapes the bundle")
    return resolved


def cleanup_spec() -> dict[str, Any]:
    contract = load(ROOT / "lifecycle-contract.json")
    require(isinstance(contract, dict) and contract.get("change") == 444, "invalid lifecycle contract")
    cleanup = contract.get("cleanup")
    require(isinstance(cleanup, dict), "lifecycle contract cleanup section is missing")
    return cleanup


def bound_paths(cleanup: dict[str, Any]) -> tuple[tuple[Path, ...], tuple[Path, ...]]:
    raw_paths = cleanup.get("paths")
    raw_preserved = cleanup.get("preserved_paths")
    require(isinstance(raw_paths, list) and raw_paths, "cleanup.paths must be an explicit nonempty list")
    require(isinstance(raw_preserved, list), "cleanup.preserved_paths must be a list")
    prefix = cleanup.get("prefix", PREFIX)
    require(isinstance(prefix, str) and prefix, "cleanup.prefix is invalid")
    paths: list[Path] = []
    for raw in raw_paths:
        target = Path(raw)
        require(target.is_absolute() and target.parent == TMP and target.name.startswith(prefix), f"unsafe cleanup path: {raw}")
        require(target.name != prefix and ".." not in target.parts, f"unsafe cleanup path: {raw}")
        paths.append(target)
    if len(set(paths)) != len(paths):
        raise RuntimeError("cleanup.paths contains duplicates")
    preserved = tuple(Path(raw) for raw in raw_preserved)
    for path in preserved:
        require(path.is_absolute() and path.name and ".." not in path.parts, f"unsafe preserved path: {path}")
    return tuple(paths), preserved


def identity(path: Path) -> tuple[int, int]:
    stat = path.stat()
    return stat.st_dev, stat.st_ino


def reject_links(path: Path) -> None:
    require(not path.is_symlink(), f"refusing symlink: {path}")
    for child in path.rglob("*"):
        require(not child.is_symlink(), f"scratch tree contains symlink: {child}")


def totals(path: Path) -> tuple[int, int]:
    files = [item for item in path.rglob("*") if item.is_file()]
    require(all(not item.is_symlink() for item in files), f"scratch tree contains symlink: {path}")
    return len(files), sum(item.stat().st_size for item in files)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--precleanup-receipt", type=Path)
    args = parser.parse_args()
    try:
        cleanup = cleanup_spec()
        paths, preserved = bound_paths(cleanup)
        pre_name = args.precleanup_receipt or cleanup.get("precleanup_receipt")
        inventory_name = cleanup.get("cleanup_inventory")
        pre = bundle_path(pre_name, "cleanup.precleanup_receipt")
        output = bundle_path(inventory_name, "cleanup.cleanup_inventory")
        require(pre.is_file(), "passing precleanup lifecycle receipt is required")
        pre_row = load(pre)
        require(isinstance(pre_row, dict) and pre_row.get("status") == "pass", "precleanup receipt is not passing")
        require(not output.exists(), "cleanup inventory already exists")
        goal_raw = cleanup.get("goal_path")
        goal_sha = cleanup.get("goal_sha256")
        require(isinstance(goal_raw, str) and isinstance(goal_sha, str) and len(goal_sha) == 64, "GOAL identity is not pinned")
        goal = Path(goal_raw)
        require(goal.is_absolute() and goal.name == "GOAL.md" and goal.is_file() and not goal.is_symlink(), "pinned GOAL.md is unavailable")
        before_goal = sha(goal)
        require(before_goal == goal_sha, "GOAL.md differs from the pinned digest")
        for target in preserved:
            require(target.is_dir() and not target.is_symlink(), f"preserved target unavailable: {target}")
        before_targets = {str(path): list(identity(path)) for path in preserved}
        removed: list[dict[str, Any]] = []
        for target in paths:
            require(target.is_dir(), f"cleanup directory unavailable: {target}")
            reject_links(target)
            files, bytes_count = totals(target)
            removed.append({"path": str(target), "regular_files": files, "regular_file_bytes": bytes_count})
        for target in paths:
            shutil.rmtree(target)
            require(not target.exists(), f"cleanup did not remove {target}")
        require({str(path): list(identity(path)) for path in preserved} == before_targets, "preserved target identity changed")
        require(sha(goal) == before_goal == goal_sha, "GOAL.md changed during cleanup")
        row = {
            "schema": "litchi-0444-cleanup-inventory-v1",
            "change": 444,
            "status": "pass",
            "precleanup_receipt": str(pre.relative_to(ROOT)),
            "cleanup_paths": [str(path) for path in paths],
            "removed": removed,
            "preserved_target_directory_identity": before_targets,
            "goal_path": str(goal),
            "goal_sha256": goal_sha,
        }
        output.parent.mkdir(exist_ok=True)
        output.write_text(json.dumps(row, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        print(json.dumps({"status": "pass", "removed_directories": len(paths), "regular_file_bytes": sum(item["regular_file_bytes"] for item in removed)}, sort_keys=True))
        return 0
    except (OSError, TypeError, ValueError, RuntimeError, json.JSONDecodeError) as error:
        print(f"CLEANUP INVALID: {error}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
