#!/usr/bin/env python3
"""Remove the five 0435 scratch directories after portable proof.

The cleanup command is terminal and deliberately narrow: it requires the
precleanup portable receipt, preserves both shared Cargo target directories
by device/inode, and verifies the repository goal document byte-for-byte by
its pinned digest.  It writes its inventory only after every removal and
preservation check succeeds.
"""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
import shutil
from typing import Any

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
PRECHECK = ROOT / "checks" / "precleanup-portable.json"
OUTPUT = ROOT / "checks" / "cleanup-inventory.json"
GOAL_SHA256 = "bed4058bb76330daab8ce9d4bceff639ab3fbd7ea06634158bef41b133c4d1f1"
PATHS = tuple(
    Path("/tmp") / name
    for name in (
        "litchi-goal-0435-binaries",
        "litchi-goal-0435-odt",
        "litchi-goal-0435-odt-tests",
        "litchi-goal-0435-odt-harness",
        "litchi-goal-0435-common-tests",
    )
)
PRESERVED = (REPO / "target", REPO / "tools" / "perf-baseline" / "target")


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def require(condition: bool, message: str) -> None:
    if not condition:
        raise RuntimeError(message)


def directory_identity(path: Path) -> tuple[int, int]:
    stat = path.stat()
    return stat.st_dev, stat.st_ino


def file_totals(path: Path) -> tuple[int, int]:
    files = [
        candidate
        for candidate in path.rglob("*")
        if candidate.is_file() and not candidate.is_symlink()
    ]
    return len(files), sum(candidate.stat().st_size for candidate in files)


def main() -> int:
    precheck = json.loads(PRECHECK.read_text(encoding="utf-8"))
    require(
        precheck.get("status") == "pass" and precheck.get("exit_code") == 0,
        f"{PRECHECK}: portable precleanup gate is not passing",
    )
    require(not OUTPUT.exists(), f"{OUTPUT}: cleanup inventory already exists")

    for path in PRESERVED:
        require(path.is_dir() and not path.is_symlink(), f"preserved target is unavailable: {path}")
    before_targets = {
        str(path): list(directory_identity(path))
        for path in PRESERVED
    }
    goal = REPO / "docs" / "GOAL.md"
    goal_before = digest(goal)
    require(goal_before == GOAL_SHA256, "docs/GOAL.md digest differs from the pinned goal")

    removed: list[dict[str, Any]] = []
    for path in PATHS:
        require(path.is_dir() and not path.is_symlink(), f"cleanup directory is unavailable: {path}")
        files, bytes_count = file_totals(path)
        removed.append(
            {
                "path": str(path),
                "regular_files": files,
                "regular_file_bytes": bytes_count,
            }
        )

    for path in PATHS:
        shutil.rmtree(path)
        require(not path.exists(), f"cleanup directory remains after removal: {path}")

    after_targets = {
        str(path): list(directory_identity(path))
        for path in PRESERVED
    }
    require(after_targets == before_targets, "shared Cargo target directory identity changed")
    require(digest(goal) == goal_before == GOAL_SHA256, "docs/GOAL.md changed during cleanup")

    record = {
        "change": 435,
        "status": "pass",
        "precleanup_receipt": str(PRECHECK.relative_to(ROOT)),
        "removed": removed,
        "preserved_target_directory_identity": before_targets,
        "user_goal_sha256": GOAL_SHA256,
    }
    OUTPUT.write_text(json.dumps(record, indent=2, sort_keys=True) + "\n", encoding="utf-8")
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
