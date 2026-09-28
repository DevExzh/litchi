"""Audit only retained 0802 failed attempts without rerunning commands.

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
HELPER_PATHS = [
    "crates/litchi-opc/src/xml_attributes.rs",
    "crates/litchi-ole-common/src/xml_attributes.rs",
    "crates/litchi-sign/src/xml_attributes.rs",
    "crates/litchi-xldm/src/xml_attributes.rs",
    "crates/xml-minifier/src/xml_attributes.rs",
    "crates/litchi-opc/src/xml_attributes/tests.rs",
]


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


def failed_patches() -> dict[str, dict[str, Any]]:
    result = {}
    for path in sorted(p.glob("*-failed-*.patch")):
        result[path.name] = artifact(path)
    return result


def archive_name(source_name: str) -> str:
    if source_name.endswith("crates/litchi-opc/src/xml_attributes/tests.rs"):
        return "litchi-opc-xml_attributes-tests.rs"
    return Path(source_name).parent.parent.name + "-xml_attributes.rs"


def patch_paths(path: Path) -> tuple[set[str], set[str]]:
    old = {line.split()[1] for line in path.read_text().splitlines()
           if line.startswith("--- ")}
    new = {line.split()[1] for line in path.read_text().splitlines()
           if line.startswith("+++ ")}
    return old, new


def audit_quality_failure(root: Path, records: dict[str, Any]) -> None:
    """Check the known retained Clippy attempt when that archive exists."""

    relocation_path = root / "relocation.json"
    relocation = c.read(relocation_path)
    assert set(relocation) == {"candidate", "candidate.patch", "fix",
                               "quality", "reason", "test-src"}
    assert relocation["quality"] == root.name
    assert relocation["candidate"] == "candidate-failed-0"
    assert relocation["test-src"] == "test-src-failed-0"
    reason = str(relocation["reason"]).lower()
    assert "clippy" in reason and "question_mark" in reason

    candidate_root = p / relocation["candidate"]
    test_root = p / relocation["test-src"]
    archive_inputs = c.read(root / "archive-inputs.json")
    assert tree(candidate_root) == archive_inputs
    assert tree(test_root)
    baseline = c.read(p / "source.json")["files"]
    for name in HELPER_PATHS:
        archived = candidate_root / ("before" if name else "") / archive_name(name)
        failed_test = test_root / "before" / name
        assert c.sha(archived) == baseline[name]
        assert c.sha(failed_test) == c.sha(archived)
        failed_after = test_root / "after" / name
        assert c.sha(failed_after) == c.sha(
            candidate_root / "after" / archive_name(name)
        )

    patch = p / relocation["candidate.patch"]
    assert patch.is_file() and not patch.is_symlink()
    old, new = patch_paths(patch)
    assert old == {f"a/{name}" for name in HELPER_PATHS}
    assert new == {f"b/{name}" for name in HELPER_PATHS}

    rows = c.read(root / "receipts.json")
    assert isinstance(rows, list) and len(rows) == 6
    assert [row["exit_code"] for row in rows[:5]] == [0] * 5
    assert rows[-1]["exit_code"] != 0
    assert "clippy::question_mark" in (
        root / Path(rows[-1]["log"]["path"]).name
    ).read_text(errors="replace")
    records["quality_failure"] = {
        "marker": "clippy::question_mark",
        "stage": "after helper Clippy",
        "relocation": artifact(relocation_path),
        "candidate_archive": [
            artifact(candidate_root / name) for name in sorted(archive_inputs)
        ],
        "test_source": [artifact(test_root / name) for name in sorted(tree(test_root))],
        "patch": artifact(patch),
    }


def result() -> dict[str, Any]:
    attempts = [audit_attempt(root) for root in attempt_roots()]
    for root, records in zip(attempt_roots(), attempts):
        if root.name.startswith("quality-failed-"):
            audit_quality_failure(root, records)
    cleanup = p / "cleanup.json"
    if cleanup.is_file():
        value = c.read(cleanup)
        assert isinstance(value, dict)
        assert isinstance(value.get("removed_failed_binaries"), list)
    return {
        "schema": "litchi.performance.0802.failure-audit.v1",
        "production_adoption": False,
        "production_source_changed": False,
        "failed_attempts": attempts,
        "patches": failed_patches(),
        "path_policy": "only top-level retained failed-attempt roots are audited; no failed command is rerun",
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
    print(f"0802 failure audit PASS: {len(expected['failed_attempts'])} retained attempts")
