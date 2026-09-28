"""Audit only retained 0804 failed attempts without rerunning commands.

The packet is allowed to contain no failed attempt at all. If root retains a
failed build or quality attempt, this reader records the files and receipt
metadata that are actually present; it does not invent a failure, assume a
particular compiler diagnostic, or rerun any command.
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Any

import custody as c


p = c.P
FAILED_ROOT = re.compile(r"^[a-z][a-z0-9-]*-failed-[0-9]+$")
COMPONENT_ROOTS = {"candidate", "probe-src", "test-src"}


def artifact(path: Path) -> dict[str, Any]:
    path = path.resolve()
    assert path.is_file() and not path.is_symlink(), path
    return {
        "path": str(path.relative_to(p.resolve())),
        "bytes": path.stat().st_size,
        "sha256": c.sha(path),
    }


def tree(root: Path) -> dict[str, str]:
    root = root.resolve()
    return {
        str(path.relative_to(root)): c.sha(path)
        for path in sorted(root.rglob("*"))
        if path.is_file() and not path.is_symlink()
    }


def retained_log(root: Path, row: dict[str, Any]) -> dict[str, Any]:
    descriptor = row.get("log")
    assert isinstance(descriptor, dict), row
    original = Path(descriptor["path"])
    assert original.is_absolute(), descriptor
    retained = root / original.name
    expected = dict(descriptor)
    expected["path"] = str(retained.resolve().relative_to(p.resolve()))
    actual = artifact(retained)
    assert actual == expected, (retained, actual, expected)
    return actual


def receipts(root: Path, name: str) -> dict[str, Any] | None:
    path = root / name
    if not path.is_file():
        return None
    rows = c.read(path)
    assert isinstance(rows, list), path
    previous = None
    logs = []
    for index, row in enumerate(rows):
        assert isinstance(row, dict), (path, index)
        started = row.get("started")
        ended = row.get("ended")
        assert isinstance(started, (int, float))
        assert isinstance(ended, (int, float))
        assert started <= ended
        if previous is not None:
            assert previous <= started
        previous = ended
        assert isinstance(row.get("exit_code"), int)
        logs.append(retained_log(root, row))
    return {
        "path": str(path.resolve().relative_to(p.resolve())),
        "bytes": path.stat().st_size,
        "sha256": c.sha(path),
        "commands": [row.get("command") for row in rows],
        "exit_codes": [row["exit_code"] for row in rows],
        "logs": logs,
    }


def attempt_roots() -> list[Path]:
    result = []
    for path in sorted(p.iterdir()):
        if not path.is_dir() or path.name in COMPONENT_ROOTS:
            continue
        if FAILED_ROOT.fullmatch(path.name) is None:
            continue
        # A top-level failure root must contain a relocation, command, or
        # receipt record. Component archives are intentionally ignored.
        if any((path / name).is_file() for name in
               ("relocation.json", "commands.json", "receipts.json")):
            result.append(path)
    return result


def audit_attempt(root: Path) -> dict[str, Any]:
    files = tree(root)
    records: dict[str, Any] = {
        "id": root.name,
        "files": [artifact(root / name) for name in sorted(files)],
    }
    relocation = root / "relocation.json"
    if relocation.is_file():
        value = c.read(relocation)
        assert isinstance(value, dict), relocation
        records["relocation"] = artifact(relocation)
        records["relocation_keys"] = sorted(value)
    receipt = receipts(root, "commands.json")
    if receipt is None:
        receipt = receipts(root, "receipts.json")
    if receipt is not None:
        records["receipts"] = receipt
    return records


def retained_artifact(root: Path, descriptor: dict[str, Any]) -> dict[str, Any]:
    """Verify a failed receipt's artifact after relocation into its archive."""

    assert isinstance(descriptor, dict)
    retained = root / Path(descriptor["path"]).name
    actual = artifact(retained)
    assert actual["bytes"] == descriptor["bytes"]
    assert actual["sha256"] == descriptor["sha256"]
    return actual


def audit_profile_capture_failure(
    root: Path, records: dict[str, Any], build_binary: dict[str, Any]
) -> None:
    """Audit the one-child owner-resolution failure retained by the driver."""

    relocation_path = root / "relocation.json"
    relocation = c.read(relocation_path)
    assert set(relocation) == {
        "child_exit_codes", "driver_exit_code", "original", "reason",
        "repair", "retained",
    }
    assert relocation["original"] == "profiles"
    assert relocation["retained"] == root.name
    assert relocation["child_exit_codes"] == [0]
    assert relocation["driver_exit_code"] == 1
    reason = str(relocation["reason"]).lower()
    assert "before_construct" in reason
    assert "termination" in reason

    rows = c.read(root / "receipts.json")
    assert isinstance(rows, list) and len(rows) == 1
    row = rows[0]
    assert row["exit_code"] == 0
    assert row["binary"] == build_binary
    assert set(row["artifacts"]) == {
        "0-distinct-0-construct-before.callgrind",
        "0-distinct-0-construct-before.json",
        "0-distinct-0-construct-before.log",
    }
    retained = {
        name: retained_artifact(root, descriptor)
        for name, descriptor in row["artifacts"].items()
    }
    assert retained["0-distinct-0-construct-before.log"] == retained_artifact(
        root, row["log"]
    )
    assert retained["0-distinct-0-construct-before.json"] == retained_artifact(
        root, row["report"]
    )
    assert not list(root.glob("*.callgrind.1"))
    assert not list(root.glob("*.callgrind.2"))
    callgrind = (
        root / "0-distinct-0-construct-before.callgrind"
    ).read_text(encoding="utf-8", errors="replace")
    assert "part: 1" in callgrind
    assert "Trigger: Program termination" in callgrind
    assert "summary: 0" in callgrind
    assert "totals: 0" in callgrind
    records["profile_capture_failure"] = {
        "driver_exit_code": relocation["driver_exit_code"],
        "child_exit_codes": relocation["child_exit_codes"],
        "relocation": artifact(relocation_path),
        "binary": dict(row["binary"]),
        "artifacts": retained,
        "positive_dump_present": False,
        "termination_dump_zero": True,
    }


def failed_patches() -> dict[str, dict[str, Any]]:
    result = {}
    for path in sorted(p.glob("*-failed-*.patch")):
        result[path.name] = artifact(path)
    return result


def result() -> dict[str, Any]:
    roots = attempt_roots()
    attempts = [audit_attempt(root) for root in roots]
    build_binary = c.read(p / "build/build.json")["binary"]
    for root, records in zip(roots, attempts):
        if root.name == "profiles-failed-0":
            audit_profile_capture_failure(root, records, build_binary)
    cleanup = p / "cleanup.json"
    if cleanup.is_file():
        value = c.read(cleanup)
        assert isinstance(value, dict)
        assert isinstance(value.get("removed_failed_binaries"), list)
    return {
        "schema": "litchi.performance.0804.failure-audit.v1",
        "production_adoption": False,
        "production_source_changed": False,
        "failed_attempts": attempts,
        "patches": failed_patches(),
        "path_policy": (
            "only top-level retained failed-attempt roots are audited; "
            "no failed command is rerun; an empty list is the expected clean-run result"
        ),
    }


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--write", action="store_true")
    group.add_argument("--check", action="store_true")
    args = parser.parse_args()
    output = p / "failure-audit.json"
    expected = result()
    if args.write:
        assert not output.exists(), output
        c.write(output, expected)
    else:
        assert c.read(output) == expected
    print(f"0804 failure audit PASS: {len(expected['failed_attempts'])} retained attempts")
