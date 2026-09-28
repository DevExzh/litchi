"""Validate the complete 0802 evidence packet without running a benchmark."""

from __future__ import annotations

import argparse
import re
import subprocess
import sys
from pathlib import Path

import custody as c
from fixture_check import check


p = c.P
ROOT = c.ROOT
NATIVE_CHILDREN = 936
PROFILE_CHILDREN = 312
NATIVE_SAMPLES = 28_080
TOTAL_REPORTS = 1_248
TOTAL_SAMPLES = 28_392
HELPERS = [
    "litchi-opc",
    "litchi-ole-common",
    "litchi-sign",
    "litchi-xldm",
    "xml-minifier",
]


def packet_path(value: str | Path) -> Path:
    path = Path(value)
    if path.is_absolute():
        return path.resolve()
    for candidate in (p / path, ROOT / path):
        if candidate.exists():
            return candidate.resolve()
    return (p / path).resolve()


def artifact(value: dict) -> Path:
    assert isinstance(value, dict), value
    path = packet_path(value["path"])
    assert path.is_file() and not path.is_symlink(), path
    assert c.artifact(path) == {
        "path": str(path),
        "bytes": value["bytes"],
        "sha256": value["sha256"],
    }, path
    return path


def check_descriptor(value: dict, label: str) -> None:
    try:
        artifact(value)
    except (AssertionError, KeyError) as error:
        raise AssertionError(f"{label}: {error}") from error


def check_external(value: dict, label: str, cleanup: dict | None) -> None:
    assert isinstance(value, dict), label
    path = Path(value["path"])
    assert path.is_absolute(), label
    if path.is_file() and not path.is_symlink():
        assert c.artifact(path) == value, label
        return
    assert cleanup is not None, (label, path)
    assert cleanup["target"] == str(c.TARGET)
    assert cleanup["target_removed"] is True
    assert not c.TARGET.exists()
    removed = [*cleanup.get("removed_binaries", []),
               *cleanup.get("removed_failed_binaries", [])]
    assert value in removed, label


def monotonic_receipts(rows: list[dict], label: str) -> None:
    previous = None
    for index, row in enumerate(rows):
        assert row["started"] <= row["ended"], (label, index)
        if previous is not None:
            assert previous <= row["started"], (label, index)
        previous = row["ended"]
        assert row["exit_code"] == 0, (label, index, row["exit_code"])
        check_descriptor(row["log"], f"{label}[{index}].log")


def source_and_environment() -> tuple[dict, dict, dict]:
    origin = c.read(p / "origin.json")
    source = c.read(p / "source.json")
    current = c.source()
    assert source["revision"] == origin["base"]
    # The packet may be checked before or after its own commit. The commit
    # identity is intentionally open; production file hashes and ancestry are
    # the authoritative post-commit checks.
    assert current["files"] == source["files"]
    subprocess.run(
        ["git", "merge-base", "--is-ancestor", origin["base"], "HEAD"],
        cwd=ROOT,
        check=True,
    )
    worktrees = subprocess.check_output(
        ["git", "worktree", "list", "--porcelain"], cwd=ROOT, text=True
    )
    recorded = origin["worktrees"].strip().split("\n\n")
    actual = worktrees.strip().split("\n\n")
    recorded_root = next(block for block in recorded if block.startswith(f"worktree {ROOT}\n"))
    actual_root = next(block for block in actual if block.startswith(f"worktree {ROOT}\n"))
    without_head = lambda block: [
        line for line in block.splitlines() if not line.startswith("HEAD ")
    ]
    assert without_head(actual_root) == without_head(recorded_root)
    assert [block for block in actual if block != actual_root] == [
        block for block in recorded if block != recorded_root
    ]
    for name, digest in origin["unrelated"].items():
        assert c.sha(ROOT / name) == digest, name
    for name, digest in c.read(p / "architecture-inputs.json").items():
        assert c.sha(ROOT / name) == digest, name
    assert c.sha(p / "workspace-Cargo.lock") == c.sha(ROOT / "Cargo.lock")
    return origin, source, current


def check_inheritance(source: dict) -> None:
    inheritance = c.read(p / "inheritance.json")
    for name in (
        "baseline_helper",
        "candidate_helper",
        "candidate_manifest",
        "previous_seal",
        "parser",
        "preflight_seal",
    ):
        check_descriptor(inheritance[name], f"inheritance.{name}")
    assert c.sha(p / "probe-src/src/baseline.rs") == inheritance["baseline_helper"]["sha256"]
    assert c.sha(p / "probe-src/src/candidate.rs") == inheritance["candidate_helper"]["sha256"]
    assert c.sha(p / "callgrind_parser.py") == inheritance["parser"]["sha256"]
    assert source["files"]["crates/litchi-opc/src/xml_attributes.rs"] == inheritance[
        "baseline_helper"
    ]["sha256"]


def check_build(source: dict, cleanup: dict | None) -> dict:
    root = p / "build"
    build = c.read(root / "build.json")
    inputs = c.read(artifact(build["inputs"]))
    for name, digest in inputs["probe"].items():
        assert c.sha(p / "probe-src" / name) == digest, name
    for name, digest in inputs["frozen"].items():
        assert c.sha(packet_path(name)) == digest, name
    rows = build["rows"]
    assert len(rows) == 6
    monotonic_receipts(rows, "build")
    assert build["environment"] == {
        "CARGO_TARGET_DIR": str(c.TARGET),
        "CARGO_BUILD_JOBS": "2",
        "CARGO_INCREMENTAL": "0",
        "RUSTFLAGS": None,
    }
    check_external(build["binary"], "build.binary", cleanup)
    assert c.read(root / "commands.json") == rows
    assert c.read(root / "inputs.json") == inputs
    assert c.artifact(p / "probe-src/Cargo.lock") == build["lock"]
    assert c.source()["files"] == source["files"]
    return build


def check_quality(source: dict) -> None:
    root = p / "quality"
    archive_inputs = c.read(root / "archive-inputs.json")
    for name, digest in archive_inputs.items():
        assert c.sha(p / "candidate" / name) == digest, name
    complete = c.read(root / "complete.json")
    receipts = c.read(root / "receipts.json")
    assert complete["rows"] == receipts and len(receipts) == 6
    monotonic_receipts(receipts, "quality")
    for leg in ("before", "after"):
        text = (root / f"{leg}-1.log").read_text()
        rows = re.findall(r"test result: ok\. (\d+) passed; (\d+) failed; (\d+) ignored; (\d+) measured; (\d+) filtered out", text)
        assert rows
        totals = [sum(int(row[i]) for row in rows) for i in range(5)]
        assert totals[0] > 0
        assert totals[1:] == [0, 0, 0, 0]
    for leg, files in complete["inputs"].items():
        for name, digest in files.items():
            assert c.sha(p / "test-src" / leg / name) == digest, (leg, name)
    for owner in HELPERS:
        before = p / "candidate/before" / f"{owner}-xml_attributes.rs"
        after = p / "candidate/after" / f"{owner}-xml_attributes.rs"
        assert before.read_bytes() == (ROOT / f"crates/{owner}/src/xml_attributes.rs").read_bytes()
        assert c.sha(p / "test-src/before" / f"crates/{owner}/src/xml_attributes.rs") == c.sha(before)
        assert c.sha(p / "test-src/after" / f"crates/{owner}/src/xml_attributes.rs") == c.sha(after)
    before_tests = p / "candidate/before/litchi-opc-xml_attributes-tests.rs"
    after_tests = p / "candidate/after/litchi-opc-xml_attributes-tests.rs"
    assert before_tests.read_bytes() == (ROOT / "crates/litchi-opc/src/xml_attributes/tests.rs").read_bytes()
    assert c.sha(p / "test-src/before/crates/litchi-opc/src/xml_attributes/tests.rs") == c.sha(before_tests)
    assert c.sha(p / "test-src/after/crates/litchi-opc/src/xml_attributes/tests.rs") == c.sha(after_tests)
    saved_source = complete.get("source")
    if saved_source is not None:
        check_descriptor(saved_source, "quality.source")
        assert c.read(packet_path(saved_source["path"]))["files"] == source["files"]


def check_lane(
    lane: str, plan: dict, source: dict, build: dict, cleanup: dict | None
) -> tuple[int, int]:
    root = p / lane
    complete = c.read(root / "complete.json")
    receipts = c.read(artifact(complete["receipts"]))
    source_descriptor = complete["source"]
    check_descriptor(source_descriptor, f"{lane}.source")
    assert c.read(packet_path(source_descriptor["path"])) == source
    expected = len(c.read(p / "cases.json")) * len(plan[lane]["orders"]) * 2 * 2
    assert complete["children"] == len(receipts) == expected
    settings = plan[lane]
    ids = [row["id"] for row in c.read(p / "cases.json")]
    schedule = [
        (block, case, mode, leg)
        for block, order in enumerate(settings["orders"])
        for case in ids
        for mode in ("construct", "consume")
        for leg in order
    ]
    assert len(schedule) == expected
    previous = None
    for index, (row, expected_identity) in enumerate(zip(receipts, schedule)):
        coordinate = (
            row.get("repeat", row.get("block"))
            if lane == "profiles"
            else row.get("block")
        )
        assert (coordinate, row["case"], row["mode"], row["leg"]) == expected_identity
        assert row["binary"]["sha256"] == build["binary"]["sha256"]
        assert row["binary"]["bytes"] == build["binary"]["bytes"]
        assert row["started"] <= row["ended"]
        if previous is not None:
            assert previous <= row["started"], (lane, index)
        previous = row["ended"]
        assert row["exit_code"] == 0
        check_descriptor(row["log"], f"{lane}[{index}].log")
        check_descriptor(row["report"], f"{lane}[{index}].report")
        if lane == "profiles":
            for descriptor in row["artifacts"].values():
                check_descriptor(descriptor, f"{lane}[{index}].artifact")
    check_external(build["binary"], "capture binary", cleanup)
    return len(receipts), sum(settings["samples"] for _ in receipts)


def run_offline_audits() -> None:
    for script in (
        "failure_audit.py",
        "analyze.py",
        "root_cg_totals.py",
        "root_native_audit.py",
        "decision.py",
    ):
        subprocess.run(
            [sys.executable, "-B", str(p / script), "--check"],
            cwd=ROOT,
            check=True,
        )
    analysis = c.read(p / "analysis.json")
    paired = analysis["native"]["analysis"]["paired_by_case_mode"]
    audit = c.read(p / "root-native-audit.json")
    assert audit["reports"] == TOTAL_REPORTS
    for row in audit["rows"]:
        other = paired[f"{row['case']}/{row['mode']}"]
        assert row["ratio"] == other["ratio_median"]
        assert row["ci_low"] == other["bootstrap"]["ci_low"]
        assert row["ci_high"] == other["bootstrap"]["ci_high"]
        assert row["regression"] == other["diagnostic_regression"]


def validate(require_final_seal: bool) -> None:
    _, source, _ = source_and_environment()
    check_inheritance(source)
    cleanup = c.read(p / "cleanup.json") if (p / "cleanup.json").is_file() else None
    build = check_build(source, cleanup)
    check_quality(source)
    check()
    plan = c.read(p / "plan.json")
    native_reports, native_samples = check_lane("native", plan, source, build, cleanup)
    profile_reports, profile_samples = check_lane("profiles", plan, source, build, cleanup)
    assert native_reports + profile_reports == TOTAL_REPORTS
    assert native_samples == NATIVE_SAMPLES
    assert native_samples + profile_samples == TOTAL_SAMPLES
    run_offline_audits()
    assert not list(p.rglob("__pycache__"))
    if require_final_seal:
        subprocess.run(
            [sys.executable, "-B", str(p / "seal_packet.py")], cwd=ROOT, check=True
        )
    print(
        {
            "reports": native_reports + profile_reports,
            "samples": native_samples + profile_samples,
            "production_changed": 0,
            "final_seal_checked": require_final_seal,
        }
    )


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--require-final-seal", action="store_true")
    arguments = parser.parse_args()
    validate(arguments.require_final_seal)
