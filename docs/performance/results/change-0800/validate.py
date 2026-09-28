"""Validate the complete 0800 correctness-fix evidence packet."""

from __future__ import annotations

import argparse
import re
import subprocess
import sys
from pathlib import Path

import custody as c


p = c.P
ROOT = c.ROOT
HELPER_PATHS = [
    "crates/litchi-ole-common/src/xml_attributes.rs",
    "crates/litchi-opc/src/xml_attributes.rs",
    "crates/litchi-opc/src/xml_attributes/tests.rs",
    "crates/litchi-sign/src/xml_attributes.rs",
    "crates/litchi-xldm/src/xml_attributes.rs",
    "crates/xml-minifier/src/xml_attributes.rs",
]


def packet_path(value: str | Path) -> Path:
    path = Path(value)
    if path.is_absolute():
        return path.resolve()
    for candidate in (p / path, ROOT / path):
        if candidate.exists():
            return candidate.resolve()
    return (p / path).resolve()


def artifact(path: Path) -> dict:
    path = path.resolve()
    assert path.is_file() and not path.is_symlink(), path
    return {
        "path": str(path),
        "bytes": path.stat().st_size,
        "sha256": c.sha(path),
    }


def descriptor_matches(path: Path, expected: dict) -> None:
    actual = artifact(path)
    assert actual["bytes"] == expected["bytes"]
    assert actual["sha256"] == expected["sha256"]


def archive_name(source_name: str) -> str:
    if source_name.endswith("crates/litchi-opc/src/xml_attributes/tests.rs"):
        return "litchi-opc-xml_attributes-tests.rs"
    return Path(source_name).parent.parent.name + "-xml_attributes.rs"


def validate_patch(path: Path) -> None:
    lines = path.read_text().splitlines()
    old = {line.split()[1] for line in lines if line.startswith("--- ")}
    new = {line.split()[1] for line in lines if line.startswith("+++ ")}
    assert old == {f"a/{name}" for name in HELPER_PATHS}
    assert new == {f"b/{name}" for name in HELPER_PATHS}


def validate_origin_and_source() -> tuple[dict, dict, dict]:
    origin = c.read(p / "origin.json")
    baseline = c.read(p / "source.json")
    current = c.source()
    assert origin["base"] == baseline["revision"]
    assert subprocess.run(
        ["git", "merge-base", "--is-ancestor", origin["base"], current["revision"]],
        cwd=ROOT,
        check=False,
    ).returncode == 0
    worktrees = subprocess.check_output(
        ["git", "worktree", "list", "--porcelain"], cwd=ROOT, text=True
    )
    recorded_blocks = origin["worktrees"].strip().split("\n\n")
    current_blocks = worktrees.strip().split("\n\n")
    recorded_root = next(block for block in recorded_blocks if block.startswith(f"worktree {ROOT}\n"))
    current_root = next(block for block in current_blocks if block.startswith(f"worktree {ROOT}\n"))
    recorded_root_lines = [line for line in recorded_root.splitlines() if not line.startswith("HEAD ")]
    current_root_lines = [line for line in current_root.splitlines() if not line.startswith("HEAD ")]
    assert current_root_lines == recorded_root_lines
    assert [block for block in current_blocks if block != current_root] == [
        block for block in recorded_blocks if block != recorded_root
    ]
    for name, digest in origin["unrelated"].items():
        assert c.sha(ROOT / name) == digest, name
    for name, digest in c.read(p / "architecture-inputs.json").items():
        assert c.sha(ROOT / name) == digest, name
    assert set(current["files"]) == set(baseline["files"])
    changed = {name for name in current["files"] if current["files"][name] != baseline["files"][name]}
    manifest = c.read(p / "correction/manifest.json")
    assert changed == set(manifest["changed"]) == set(HELPER_PATHS)
    for name, digest in baseline["files"].items():
        if name not in changed:
            assert current["files"][name] == digest, name
    return origin, baseline, current


def validate_correction(baseline: dict, current: dict) -> dict:
    manifest_path = p / "correction/manifest.json"
    manifest = c.read(manifest_path)
    assert manifest["schema"] == "litchi.performance.0800.correctness.v1"
    assert manifest["base"] == baseline["revision"]
    assert manifest["performance_claim"] is False
    assert set(manifest["changed"]) == set(HELPER_PATHS)
    names = {"manifest.json"}
    for name, parts in manifest["changed"].items():
        before = packet_path(parts["before"]["path"])
        after = packet_path(parts["after"]["path"])
        descriptor_matches(before, parts["before"])
        descriptor_matches(after, parts["after"])
        assert baseline["files"][name] == parts["before"]["sha256"]
        assert current["files"][name] == parts["after"]["sha256"]
        names.add(f"before/{before.name}")
        names.add(f"after/{after.name}")
    actual = {
        str(path.relative_to(p / "correction"))
        for path in (p / "correction").rglob("*")
        if path.is_file()
    }
    assert actual == names
    validate_patch(p / "correction.patch")
    corrected_text = (p / "correction/after/litchi-opc-xml_attributes.rs").read_text()
    assert "consume the first byte" in corrected_text
    test_text = (p / "correction/after/litchi-opc-xml_attributes-tests.rs").read_text()
    assert "=long_name" in test_text and "==n0" not in test_text
    return manifest


def validate_deferred_candidate(baseline: dict, current: dict) -> None:
    root = p / "deferred-candidate"
    manifest = c.read(root / "manifest.json")
    assert manifest["schema"] == "litchi.performance.0800.candidate-manifest.v1"
    assert manifest["base_commit"] == baseline["revision"]
    assert manifest["status"] == "untested_deferred_requires_rebase"
    assert manifest["production_adoption"] is False
    assert manifest["production_source_changed"] is False
    assert "no build or measurement approval" in manifest["deferred_reason"]
    assert len(manifest["before"]) == len(manifest["after"]) == 6
    before_names = {f"before/{archive_name(name)}" for name in HELPER_PATHS}
    after_names = {f"after/{archive_name(name)}" for name in HELPER_PATHS}
    assert {
        str(path.relative_to(root))
        for path in (root / "before").iterdir()
        if path.is_file()
    } == before_names
    assert {
        str(path.relative_to(root))
        for path in (root / "after").iterdir()
        if path.is_file()
    } == after_names
    for name in HELPER_PATHS:
        assert c.sha(root / "before" / archive_name(name)) == baseline["files"][name]
        assert c.sha(root / "after" / archive_name(name)) != current["files"][name]
    design = (root / "design.md").read_text().lower()
    assert "source-only" in design and "no production edit" in design
    assert "build receipt" in design
    validate_patch(p / "deferred-candidate.patch")


def validate_quality(baseline: dict, current: dict) -> None:
    quality = p / "quality"
    inputs = c.read(quality / "inputs.json")
    for name, digest in inputs.items():
        assert c.sha(packet_path(name)) == digest, name
    saved_source = c.read(quality / "source.json")
    assert saved_source["revision"] == baseline["revision"]
    assert saved_source["files"] == current["files"]
    receipts = c.read(quality / "receipts.json")
    complete = c.read(quality / "complete.json")
    assert complete["rows"] == receipts
    assert complete["scope"].endswith("no performance claim")
    assert complete["environment"] == {
        "CARGO_BUILD_JOBS": "2",
        "CARGO_INCREMENTAL": "0",
        "CARGO_TARGET_DIR": str(c.TARGET / "full"),
        "RUSTFLAGS": None,
    }
    changed = list(c.read(p / "correction/manifest.json")["changed"])
    expected_commands = [
        [
            "rustfmt",
            "--check",
            "--edition",
            "2024",
            "--config",
            "skip_children=true",
            *[str(ROOT / name) for name in changed],
        ],
        [
            "cargo",
            "run",
            "--offline",
            "--locked",
            "--manifest-path",
            str(p / "edge-probe/Cargo.toml"),
        ],
        [
            "cargo",
            "test",
            "--offline",
            "--locked",
            "-p",
            "litchi-opc",
            "-p",
            "litchi-ole-common",
            "-p",
            "litchi-sign",
            "-p",
            "litchi-xldm",
            "-p",
            "xml-minifier",
            "--",
            "--test-threads=2",
        ],
        [
            "cargo",
            "clippy",
            "--offline",
            "--locked",
            "-p",
            "litchi-opc",
            "-p",
            "litchi-ole-common",
            "-p",
            "litchi-sign",
            "-p",
            "litchi-xldm",
            "-p",
            "xml-minifier",
            "--all-targets",
            "--",
            "-D",
            "warnings",
        ],
    ]
    assert len(receipts) == len(expected_commands) == 4
    previous_end = None
    for row, command in zip(receipts, expected_commands):
        assert row["command"] == command
        assert row["exit_code"] == 0
        assert row["started"] <= row["ended"]
        if previous_end is not None:
            assert previous_end <= row["started"]
        previous_end = row["ended"]
        descriptor_matches(packet_path(row["log"]["path"]), row["log"])
    descriptor_matches(packet_path(complete["inputs"]["path"]), complete["inputs"])
    descriptor_matches(packet_path(complete["source"]["path"]), complete["source"])
    summary = c.read(p / "test-summary.json")
    assert summary["failed"] == 0
    assert summary["filtered"] == 0
    assert summary["ignored"] == 3
    assert summary["measured"] == 0
    assert summary["passed"] == 1538
    assert summary["suites"] == 67
    assert summary["scope"] == (
        "Full default-feature tests for five affected production crates, "
        "including integration and doctest suites"
    )
    descriptor_matches(packet_path(summary["source"]["path"]), summary["source"])
    test_log = (p / "quality/2.log").read_text(errors="replace")
    assert "test result: FAILED" not in test_log
    matches = re.findall(
        r"test result: (ok|FAILED)\. (\d+) passed; (\d+) failed; "
        r"(\d+) ignored; (\d+) measured; (\d+) filtered out",
        test_log,
    )
    assert len(matches) == 67
    assert all(status == "ok" for status, *_ in matches)
    totals = [sum(int(row[index]) for row in matches) for index in range(1, 6)]
    assert totals == [1538, 0, 3, 0, 0]


def validate_failure_archive() -> None:
    subprocess.run(
        [sys.executable, "-B", str(p / "failure_audit.py"), "--check"],
        cwd=ROOT,
        check=True,
    )


def validate_binaries_and_cleanup() -> None:
    identities = c.read(p / "binary-identities.json")
    assert identities["order"] == ["initial_exploratory", "controlled_locked"]
    binaries = identities["binaries"]
    assert len(binaries) == 2
    for value in binaries:
        path = Path(value["path"])
        if path.is_file() and not path.is_symlink():
            descriptor_matches(path, value)
    cleanup_path = p / "cleanup.json"
    if cleanup_path.exists():
        cleanup = c.read(cleanup_path)
        assert set(cleanup) == {
            "removed_binaries",
            "removed_target_bytes",
            "target",
            "target_removed",
        }
        assert cleanup["target"] == str(c.TARGET)
        assert cleanup["target_removed"] is True
        assert not c.TARGET.exists()
        assert cleanup["removed_binaries"] == binaries
        assert isinstance(cleanup["removed_target_bytes"], int)
        assert cleanup["removed_target_bytes"] >= 0
    else:
        assert all(Path(value["path"]).is_file() for value in binaries)


def validate_seal(require_final_seal: bool) -> None:
    if not require_final_seal:
        return
    seal = p / "seal_packet.py"
    assert seal.is_file()
    subprocess.run([sys.executable, "-B", str(seal)], cwd=ROOT, check=True)


def validate(require_final_seal: bool) -> None:
    origin, baseline, current = validate_origin_and_source()
    validate_correction(baseline, current)
    validate_deferred_candidate(baseline, current)
    validate_quality(baseline, current)
    validate_failure_archive()
    validate_binaries_and_cleanup()
    validate_seal(require_final_seal)
    assert not list(p.rglob("__pycache__"))
    print(
        {
            "production_changed": 6,
            "quality_rows": 4,
            "test_passed": 1538,
            "binaries": 2,
            "final_seal_checked": require_final_seal,
        }
    )


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--require-final-seal", action="store_true")
    args = parser.parse_args()
    validate(args.require_final_seal)
