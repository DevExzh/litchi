"""Audit retained 0801 failures without rerunning build or quality commands."""

from __future__ import annotations

import argparse
from pathlib import Path

import custody as c


p = c.P
HELPER_PATHS = [
    "crates/litchi-opc/src/xml_attributes.rs",
    "crates/litchi-ole-common/src/xml_attributes.rs",
    "crates/litchi-sign/src/xml_attributes.rs",
    "crates/litchi-xldm/src/xml_attributes.rs",
    "crates/xml-minifier/src/xml_attributes.rs",
    "crates/litchi-opc/src/xml_attributes/tests.rs",
]


def artifact(path: Path) -> dict:
    path = path.resolve()
    assert path.is_file() and not path.is_symlink(), path
    return {
        "path": str(path.relative_to(p.resolve())),
        "bytes": path.stat().st_size,
        "sha256": c.sha(path),
    }


def descriptor_matches(path: Path, expected: dict) -> dict:
    actual = artifact(path)
    assert actual == expected, (path, actual, expected)
    return actual


def packet_descriptor(value: dict, label: str) -> dict:
    path = Path(value["path"])
    if not path.is_absolute():
        path = p / path
    assert path.is_file() and not path.is_symlink(), label
    actual = {
        "path": str(path.resolve().relative_to(p.resolve())),
        "bytes": value["bytes"],
        "sha256": value["sha256"],
    }
    assert actual == value or value["path"] == str(path.resolve()), (label, actual, value)
    return value


def tree(root: Path) -> dict[str, str]:
    root = root.resolve()
    return {
        str(path.relative_to(root)): c.sha(path)
        for path in sorted(root.rglob("*"))
        if path.is_file() and not path.is_symlink()
    }


def validate_tree(root: Path, expected: dict[str, str]) -> list[dict]:
    actual = tree(root)
    assert actual == expected, (root, actual.keys() ^ expected.keys())
    return [artifact(root / name) for name in sorted(expected)]


def archive_name(source_name: str) -> str:
    if source_name.endswith("crates/litchi-opc/src/xml_attributes/tests.rs"):
        return "litchi-opc-xml_attributes-tests.rs"
    return Path(source_name).parent.parent.name + "-xml_attributes.rs"


def retained_log(root: Path, row: dict) -> dict:
    original = Path(row["log"]["path"])
    assert original.is_absolute()
    retained = root / original.name
    expected = dict(row["log"])
    expected["path"] = str(retained.resolve().relative_to(p.resolve()))
    return descriptor_matches(retained, expected)


def receipt_audit(
    root: Path, expected_codes: list[int], marker: str, file_name: str = "receipts.json"
) -> dict:
    rows = c.read(root / file_name)
    assert len(rows) == len(expected_codes)
    previous = None
    logs = []
    for index, (row, code) in enumerate(zip(rows, expected_codes)):
        assert row["exit_code"] == code
        assert row["started"] <= row["ended"]
        if previous is not None:
            assert previous <= row["started"]
        previous = row["ended"]
        log = retained_log(root, row)
        if index == len(rows) - 1:
            assert marker in (root / Path(row["log"]["path"]).name).read_text(errors="replace")
        logs.append(log)
    return {
        "commands": [row["command"] for row in rows],
        "exit_codes": expected_codes,
        "logs": logs,
    }


def patch_artifact(path: Path) -> dict:
    old = {line.split()[1] for line in path.read_text().splitlines() if line.startswith("--- ")}
    new = {line.split()[1] for line in path.read_text().splitlines() if line.startswith("+++ ")}
    assert old == {f"a/{name}" for name in HELPER_PATHS}
    assert new == {f"b/{name}" for name in HELPER_PATHS}
    return artifact(path)


def external(value: dict, label: str, cleanup: dict | None, *, allow_replaced: bool = False) -> dict:
    path = Path(value["path"])
    assert path.is_absolute(), label
    if path.is_file() and not path.is_symlink():
        current = {"path": str(path), "bytes": path.stat().st_size, "sha256": c.sha(path)}
        if current == value:
            return value
        assert allow_replaced, (label, current, value)
        return value
    assert cleanup is not None, (label, path)
    assert cleanup["target"] == str(c.TARGET)
    assert cleanup["target_removed"] is True
    assert not c.TARGET.exists()
    removed = [*cleanup.get("removed_binaries", []), *cleanup.get("removed_failed_binaries", [])]
    assert any(item["bytes"] == value["bytes"] and item["sha256"] == value["sha256"] for item in removed), label
    return value


def audit_build_failure(cleanup: dict | None) -> dict:
    root = p / "build-failed-0"
    relocation = c.read(root / "relocation.json")
    assert set(relocation) == {"build", "probe-src", "original_binary", "binary", "reason", "fix"}
    assert relocation["build"] == root.name
    assert relocation["probe-src"] == "probe-src-failed-0"
    assert "Clippy" in relocation["reason"]
    assert "dead_code" in relocation["reason"] or "unused" in relocation["reason"]
    inputs = c.read(root / "inputs.json")
    for name, digest in inputs["frozen"].items():
        assert c.sha(p / name) == digest, name
    probe_root = p / relocation["probe-src"]
    assert tree(probe_root) == inputs["probe"]
    commands = c.read(root / "commands.json")
    assert len(commands) == 4
    receipt = receipt_audit(root, [0, 0, 0, 101], "unchecked_attributes", "commands.json")
    assert commands == c.read(root / "commands.json")
    retained = external(relocation["binary"], "failed binary", cleanup)
    original = external(relocation["original_binary"], "original binary", cleanup, allow_replaced=True)
    assert original["bytes"] == retained["bytes"]
    assert original["sha256"] == retained["sha256"]
    return {
        "id": root.name,
        "stage": "standalone probe build",
        "reason": relocation["reason"],
        "fix": relocation["fix"],
        "relocation": artifact(root / "relocation.json"),
        "inputs": artifact(root / "inputs.json"),
        "probe_source": [artifact(probe_root / name) for name in sorted(inputs["probe"])],
        "receipts": receipt,
        "failed_binary": retained,
        "original_binary": original,
    }


def audit_quality_failure() -> dict:
    root = p / "quality-failed-0"
    relocation = c.read(root / "relocation.json")
    assert set(relocation) == {"candidate", "candidate.patch", "fix", "quality", "reason", "test-src"}
    assert relocation["quality"] == root.name
    assert relocation["candidate"] == "candidate-failed-0"
    assert relocation["test-src"] == "test-src-failed-0"
    assert "Clippy" in relocation["reason"]
    assert "unnecessary_map_or" in relocation["reason"] and "map_identity" in relocation["reason"]
    archive_inputs = c.read(root / "archive-inputs.json")
    candidate_root = p / relocation["candidate"]
    assert tree(candidate_root) == archive_inputs
    test_root = p / relocation["test-src"]
    test_files = tree(test_root)
    assert test_files
    baseline = c.read(p / "source.json")["files"]
    for name in HELPER_PATHS:
        archived = candidate_root / "before" / archive_name(name)
        failed_test = test_root / "before" / name
        assert c.sha(archived) == baseline[name]
        assert c.sha(failed_test) == c.sha(archived)
        failed_after = test_root / "after" / name
        assert c.sha(failed_after) == c.sha(candidate_root / "after" / archive_name(name))
    receipt = receipt_audit(root, [0, 0, 0, 0, 0, 101], "unnecessary_map_or")
    patch = patch_artifact(p / relocation["candidate.patch"])
    return {
        "id": root.name,
        "stage": "isolated helper quality",
        "reason": relocation["reason"],
        "fix": relocation["fix"],
        "relocation": artifact(root / "relocation.json"),
        "archive_inputs": artifact(root / "archive-inputs.json"),
        "candidate_archive": [artifact(candidate_root / name) for name in sorted(archive_inputs)],
        "test_source": [artifact(test_root / name) for name in sorted(test_files)],
        "receipts": receipt,
        "normalized_patch": patch,
    }


def result() -> dict:
    build_root = p / "build-failed-0"
    quality_root = p / "quality-failed-0"
    cleanup = c.read(p / "cleanup.json") if (p / "cleanup.json").is_file() else None
    failed = []
    if build_root.exists():
        failed.append(audit_build_failure(cleanup))
    if quality_root.exists():
        failed.append(audit_quality_failure())
    assert failed, "no retained failure archive"
    return {
        "schema": "litchi.performance.0801.failure-audit.v1",
        "production_adoption": False,
        "production_source_changed": False,
        "failed_attempts": failed,
        "patches": {
            "quality_failed_0": patch_artifact(p / "candidate-failed-0.patch")
        },
        "path_policy": "retained failed logs are resolved by archive basename; no failed command is rerun",
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
    print("0801 failure audit PASS")
