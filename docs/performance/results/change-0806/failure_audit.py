"""Replay the retained 0806 failed-attempt custody without rerunning work.

This reader intentionally derives its result from files that are present.
Numbered setup archives are required by this workflow, while the retained
failed-receipt list is derived from actual nonzero exit codes.  A failure is
never inferred from a missing artifact or from an earlier packet.
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Any

import custody as c


P = c.P
FAILED_ROOT = re.compile(r"^[a-z][a-z0-9-]*-failed-[0-9]+$")
COMPONENT_ROOTS = {"candidate", "probe-src", "test-src"}
SETUP_ATTEMPT = re.compile(r"^setup-attempt-([0-9]+)\.json$")
PROBE_TREE = (
    "Cargo.lock",
    "Cargo.toml",
    "Cargo.toml.template",
    "src/allocation_metrics.rs",
    "src/counting_allocator.rs",
    "src/main.rs",
)
TEST_RESULT = re.compile(
    r"^test result: (?:ok|FAILED)\.\s+(\d+) passed;\s+(\d+) failed;\s+"
    r"(\d+) ignored;\s+(\d+) measured;\s+(\d+) filtered out;",
    re.MULTILINE,
)
QUALITY_FAILURE_FILES = {"00.log", "01.log", "checks.json", "source.json"}
SHAPES = ("tiny", "medium", "large", "vendor", "unicode-vendor", "valid-4attr")
MODES = ("capture", "commit", "lifecycle")
CASES = tuple((shape, mode) for shape in SHAPES for mode in MODES)
QUALIFICATION_COUNT = len(CASES)


def artifact(path: Path) -> dict[str, Any]:
    path = path.resolve()
    assert path.is_file() and not path.is_symlink(), path
    return {
        "path": str(path.relative_to(P.resolve())),
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


def external_identity(value: dict[str, Any], label: str) -> None:
    """Check an absolute retained-binary descriptor when it still exists."""

    assert isinstance(value, dict), label
    raw = value.get("path")
    assert isinstance(raw, str) and Path(raw).is_absolute(), (label, value)
    path = Path(raw)
    if not path.is_file() or path.is_symlink():
        return
    actual = {
        "bytes": path.stat().st_size,
        "sha256": c.sha(path),
    }
    assert actual == {"bytes": value.get("bytes"), "sha256": value.get("sha256")}, (label, actual, value)


def archive_record(root: Path) -> dict[str, Any]:
    """Record a retained archive without executing anything from it."""

    assert root.is_dir() and not root.is_symlink(), root
    names = sorted(tree(root))
    return {
        "path": str(root.resolve().relative_to(P.resolve())),
        "files": [artifact(root / name) for name in names],
    }


def retained_log(root: Path, row: dict[str, Any]) -> dict[str, Any]:
    descriptor = row.get("log")
    assert isinstance(descriptor, dict), row
    original = Path(descriptor["path"])
    assert original.is_absolute(), descriptor
    retained = root / original.name
    expected = dict(descriptor)
    expected["path"] = str(retained.resolve().relative_to(P.resolve()))
    actual = artifact(retained)
    assert actual == expected, (retained, actual, expected)
    return actual


def receipt_file(root: Path, name: str) -> dict[str, Any] | None:
    path = root / name
    if not path.is_file():
        return None
    rows = c.read(path)
    assert isinstance(rows, list), path
    previous = None
    logs = []
    for index, row in enumerate(rows):
        assert isinstance(row, dict), (path, index)
        started, ended = row.get("started"), row.get("ended")
        assert isinstance(started, (int, float))
        assert isinstance(ended, (int, float)) and started <= ended
        if previous is not None:
            assert previous <= started
        previous = ended
        assert isinstance(row.get("exit_code"), int)
        logs.append(retained_log(root, row))
    assert any(code != 0 for code in [row["exit_code"] for row in rows]), (
        root, name, "failed archive has no failed receipt"
    )
    return {
        "path": str(path.resolve().relative_to(P.resolve())),
        "bytes": path.stat().st_size,
        "sha256": c.sha(path),
        "commands": [row.get("command") for row in rows],
        "exit_codes": [row["exit_code"] for row in rows],
        "logs": logs,
    }


def attempt_roots() -> list[Path]:
    roots = []
    for path in sorted(P.iterdir()):
        if not path.is_dir() or path.name in COMPONENT_ROOTS:
            continue
        if FAILED_ROOT.fullmatch(path.name) is None:
            continue
        if any((path / name).is_file()
               for name in ("relocation.json", "commands.json", "receipts.json")):
            roots.append(path)
    return roots


def audit_attempt(root: Path) -> dict[str, Any]:
    files = tree(root)
    record: dict[str, Any] = {
        "id": root.name,
        "files": [artifact(root / name) for name in sorted(files)],
    }
    relocation = root / "relocation.json"
    if relocation.is_file():
        value = c.read(relocation)
        assert isinstance(value, dict), relocation
        record["relocation"] = artifact(relocation)
        record["relocation_keys"] = sorted(value)
    receipt = receipt_file(root, "commands.json")
    if receipt is None:
        receipt = receipt_file(root, "receipts.json")
    if receipt is not None:
        record["receipts"] = receipt
        record["failed_indices"] = [
            index for index, code in enumerate(receipt["exit_codes"])
            if code != 0
        ]
        assert record["failed_indices"], root
    return record


def quality_failure() -> dict[str, Any]:
    """Audit the retained production quality stop without rerunning it."""
    root = P / "quality-0"
    assert root.is_dir() and not root.is_symlink(), root
    files = {str(path.relative_to(root)) for path in root.rglob("*")
             if path.is_file() and not path.is_symlink()}
    assert files == QUALITY_FAILURE_FILES, files
    source = c.read(root / "source.json")
    application = c.read(P / "application.json")
    assert source == application.get("source"), "quality-0 source differs from candidate application"
    rows = c.read(root / "checks.json")
    assert isinstance(rows, list) and len(rows) == 2
    previous = None
    logs = []
    expected = None
    try:
        import analyze
        expected = analyze.QUALITY_COMMANDS[:2]
    except (AttributeError, ImportError):
        pass
    for index, row in enumerate(rows):
        assert isinstance(row, dict)
        assert row.get("exit_code") == (0 if index == 0 else 101)
        assert isinstance(row.get("started"), (int, float))
        assert isinstance(row.get("ended"), (int, float))
        assert row["started"] <= row["ended"]
        if previous is not None:
            assert previous <= row["started"]
        previous = row["ended"]
        if expected is not None:
            assert row.get("command") == expected[index]
        descriptor = row.get("log")
        assert isinstance(descriptor, dict)
        path = Path(descriptor.get("path", ""))
        assert path.resolve() == (root / f"{index:02}.log").resolve()
        assert artifact(root / f"{index:02}.log") == {
            "path": f"quality-0/{index:02}.log",
            "bytes": descriptor.get("bytes"),
            "sha256": descriptor.get("sha256"),
        }
        logs.append(root / f"{index:02}.log")
    assert not logs[0].read_text(errors="replace").strip()
    failed_log = logs[1].read_text(errors="replace")
    assert "method `unchecked_attributes` is never used" in failed_log
    assert "crates/litchi-sign/src/xml_attributes.rs" in failed_log
    assert "could not compile `litchi-sign`" in failed_log
    return {
        "id": "quality-0",
        "archive": archive_record(root),
        "failed_gate": 1,
        "exit_codes": [row["exit_code"] for row in rows],
        "source": artifact(root / "source.json"),
        "diagnostic": "litchi-sign unchecked_attributes unused",
    }


def quality_failure_one() -> dict[str, Any]:
    """Audit the retained visibility regression without rerunning it."""
    root = P / "quality-1"
    assert root.is_dir() and not root.is_symlink(), root
    files = {str(path.relative_to(root)) for path in root.rglob("*")
             if path.is_file() and not path.is_symlink()}
    assert files == QUALITY_FAILURE_FILES, files
    source = c.read(root / "source.json")
    application = c.read(P / "quality-amendment-application.json")
    assert source == application.get("source"), (
        "quality-1 source differs from quality amendment application"
    )
    rows = c.read(root / "checks.json")
    assert isinstance(rows, list) and len(rows) == 2
    previous = None
    logs = []
    expected = None
    try:
        import analyze
        expected = analyze.QUALITY_COMMANDS[:2]
    except (AttributeError, ImportError):
        pass
    for index, row in enumerate(rows):
        assert isinstance(row, dict)
        assert row.get("exit_code") == (0 if index == 0 else 101)
        assert isinstance(row.get("started"), (int, float))
        assert isinstance(row.get("ended"), (int, float))
        assert row["started"] <= row["ended"]
        if previous is not None:
            assert previous <= row["started"]
        previous = row["ended"]
        if expected is not None:
            assert row.get("command") == expected[index]
        descriptor = row.get("log")
        assert isinstance(descriptor, dict)
        path = Path(descriptor.get("path", ""))
        assert path.resolve() == (root / f"{index:02}.log").resolve()
        assert artifact(root / f"{index:02}.log") == {
            "path": f"quality-1/{index:02}.log",
            "bytes": descriptor.get("bytes"),
            "sha256": descriptor.get("sha256"),
        }
        logs.append(root / f"{index:02}.log")
    assert not logs[0].read_text(errors="replace").strip()
    failed_log = logs[1].read_text(errors="replace")
    assert "error[E0603]" in failed_log
    assert "error[E0599]" in failed_log
    assert "BytesStartExt" in failed_log
    assert "checked_attributes" in failed_log
    assert "crates/litchi-ole-common/src/xml_attributes.rs" in failed_log
    assert "pub(crate) trait BytesStartExt" in failed_log
    assert "could not compile `litchi-crypto`" in failed_log
    return {
        "id": "quality-1",
        "archive": archive_record(root),
        "failed_gate": 1,
        "exit_codes": [row["exit_code"] for row in rows],
        "source": artifact(root / "source.json"),
        "diagnostic": "litchi-crypto OLE visibility E0603/E0599",
    }


def setup_attempts() -> list[tuple[int, Path]]:
    found = []
    for path in sorted(P.iterdir()):
        match = SETUP_ATTEMPT.fullmatch(path.name)
        if match is None or not path.is_file() or path.is_symlink():
            continue
        found.append((int(match.group(1)), path))
    assert [number for number, _ in found] == list(range(len(found))), found
    return found


def descriptor_bytes(value: dict[str, Any], label: str) -> tuple[int, str]:
    assert isinstance(value, dict), label
    size, digest = value.get("bytes"), value.get("sha256")
    assert isinstance(size, int) and size >= 0 and isinstance(digest, str)
    assert len(digest) == 64, (label, value)
    return size, digest


def retained_from_archive(root: Path, descriptor: dict[str, Any], label: str) -> None:
    """Check a log descriptor whose receipt retains its original path."""

    original = Path(descriptor.get("path", ""))
    assert original.is_absolute(), (label, descriptor)
    retained = root / original.name
    assert retained.is_file() and not retained.is_symlink(), (label, retained)
    size, digest = descriptor_bytes(descriptor, label)
    assert retained.stat().st_size == size and c.sha(retained) == digest, (label, retained)


def retained_artifact(root: Path, descriptor: dict[str, Any], label: str) -> Path:
    assert isinstance(descriptor, dict), label
    original = Path(descriptor.get("path", ""))
    assert original.is_absolute(), (label, descriptor)
    retained = root / original.name
    assert retained.is_file() and not retained.is_symlink(), (label, retained)
    size, digest = descriptor_bytes(descriptor, label)
    assert retained.stat().st_size == size and c.sha(retained) == digest, (label, retained)
    return retained


def setup_cleanup() -> dict[str, Any] | None:
    path = P / "setup-cleanup.json"
    if not path.is_file():
        # While the target is live, the retained descriptors can be checked
        # directly.  Once it is gone, a separate cleanup receipt is required.
        assert c.TARGET.exists(), "setup binaries disappeared without setup cleanup"
        return None
    value = c.read(path)
    assert isinstance(value, dict), path
    assert value.get("target") == str(c.TARGET)
    assert value.get("target_removed") is True and not c.TARGET.exists()
    if "schema" in value:
        assert value["schema"] == "litchi.performance.0806.setup-cleanup.v1"
    removed = value.get("removed_binaries")
    assert isinstance(removed, list) and len(removed) == 9, value
    expected_names = {
        f"before-{kind}-setup-{number}"
        for number in range(3) for kind in ("native", "allocation", "profile")
    }
    assert {item.get("path") for item in removed if isinstance(item, dict)} == {
        str(c.TARGET / name) for name in expected_names
    }
    return {"path": str(path.relative_to(P)), "bytes": path.stat().st_size,
            "sha256": c.sha(path), "value": value}


def check_setup_qualification(number: int, attempt: dict[str, Any],
                              before: dict[str, Any], allocation: dict[str, Any],
                              accepted: bool) -> dict[str, Any]:
    root = P / attempt["qualification_archive"]
    assert root.is_dir() and not root.is_symlink(), root
    complete = c.read(root / "complete.json")
    assert complete.get("children") == QUALIFICATION_COUNT
    source = retained_artifact(root, complete["source"], f"setup qualification {number}.source")
    assert c.read(source) == before
    retained_artifact(root, complete["receipts"], f"setup qualification {number}.receipts")
    rows = c.read(root / "receipts.json")
    assert isinstance(rows, list) and len(rows) == QUALIFICATION_COUNT
    mismatch = []
    for index, (row, (shape, mode)) in enumerate(zip(rows, CASES)):
        assert row.get("lane") == "qualification" and row.get("block") == 0
        assert row.get("shape") == shape and row.get("mode") == mode
        assert row.get("leg") == "before" and row.get("exit_code") == 0
        assert row.get("binary") == allocation
        assert row.get("started") <= row.get("ended")
        retained_artifact(root, row["log"], f"setup qualification {number}.{index}.log")
        report = retained_artifact(root, row["report"], f"setup qualification {number}.{index}.report")
        rss = retained_artifact(root, row["rss"], f"setup qualification {number}.{index}.rss")
        assert rss.read_text().strip().isdigit()
        value = c.read(report)
        actual_shape = shape if accepted or shape != "valid-4attr" else "valid4-attr"
        assert value.get("schema") == "litchi.pptx.public-workflow-probe-0806.v1"
        assert value.get("tool") == "public-pptx-probe-0806"
        assert value.get("shape") == actual_shape and value.get("mode") == mode
        assert value.get("samples_requested") == 1 and len(value.get("samples", [])) == 1
        command = row.get("command")
        assert ("--output" in command and "-o" in command
                and Path(command[command.index("--output") + 1]).resolve()
                == Path(row["report"]["path"]).resolve()
                and command[command.index("-o") + 1] == row["rss"]["path"]
                and allocation["path"] in command)
        verification = value["samples"][0].get("verification")
        assert verification.get("semantic_check") is True
        assert verification.get("reopened") is True
        assert verification.get("expected_text") == verification.get("actual_text")
        if actual_shape != shape:
            mismatch.append(report.name)
    if accepted:
        assert not mismatch
    else:
        witness = c.read(P / "qualification-schema-mismatch-0.json")
        assert witness.get("schema") == "litchi.performance.0806.qualification-schema-mismatch.v1"
        assert witness.get("passed") is False
        assert witness.get("capture_children_succeeded") == QUALIFICATION_COUNT
        assert isinstance(witness.get("detected_by"), str)
        assert isinstance(witness.get("explanation"), str)
        assert witness["detected_by"].strip() and witness["explanation"].strip()
        entries = witness.get("metadata_mismatches")
        assert isinstance(entries, list) and len(entries) == 3
        witness_reports = []
        for index, item in enumerate(entries):
            assert item.get("actual_serialized_shape") == "valid4-attr"
            assert item.get("expected_cli_and_plan_shape") == "valid-4attr"
            witness_reports.append(retained_artifact(root, item["report"],
                                                     f"qualification mismatch {index}.report").name)
        assert sorted(witness_reports) == sorted(mismatch)
        assert sorted(mismatch) == sorted(
            f"0-valid-4attr-{mode}-before.json" for mode in MODES
        )
    return {"reports": QUALIFICATION_COUNT, "samples": QUALIFICATION_COUNT,
            "accepted": accepted, "metadata_mismatches": len(mismatch)}


def check_setup_attempt(number: int, path: Path,
                        cleanup: dict[str, Any] | None) -> dict[str, Any]:
    value = c.read(path)
    assert isinstance(value, dict), path
    assert value.get("schema") == "litchi.performance.0806.setup-attempt.v1"
    assert value.get("build_succeeded") is True
    quality_succeeded = value.get("probe_quality_succeeded") is True
    assert value.get("probe_quality_succeeded") is quality_succeeded
    started_fields = [value[name] for name in
                      ("main_captures_started", "main_timed_captures_started")
                      if name in value]
    assert started_fields and all(flag is False for flag in started_fields)
    assert value.get("cross_lane_affected") is False
    assert isinstance(value.get("reason"), str) and value["reason"].strip()
    assert value.get("build_archive") == f"build-before-setup-{number}"
    assert value.get("probe_archive") == f"probe-src-setup-{number}"
    expected_quality_archive = (f"probe-quality-before-setup-{number}"
                                if quality_succeeded else
                                f"probe-quality-before-failed-{number}")
    assert value.get("quality_archive") == expected_quality_archive
    first_failed = value.get("first_failed_gate")
    if quality_succeeded:
        assert first_failed is None and "first_failed_gate" not in value
    else:
        assert isinstance(first_failed, int) and first_failed >= 0
    archived_source = P / "source.json"
    production = value.get("production_source")
    assert isinstance(production, dict)
    size, digest = descriptor_bytes(production, "setup production source")
    assert archived_source.stat().st_size == size and c.sha(archived_source) == digest
    assert c.read(archived_source).get("revision") == "e3ff267ee3454e71d66f177f54f3cd05e0d9cce5"

    build_root = P / value["build_archive"]
    quality_root = P / value["quality_archive"]
    probe_root = P / value["probe_archive"]
    assert build_root.is_dir() and quality_root.is_dir() and probe_root.is_dir()
    assert set(tree(probe_root)) == set(PROBE_TREE), probe_root
    setup_probe = {name: c.sha(probe_root / name) for name in PROBE_TREE}

    build = c.read(build_root / "build.json")
    assert isinstance(build, dict)
    assert build.get("source", {}).get("bytes") == size
    assert build.get("source", {}).get("sha256") == digest
    rows = c.read(build_root / "commands.json")
    assert isinstance(rows, list) and len(rows) == 3
    assert [row.get("exit_code") for row in rows] == [0, 0, 0]
    previous = None
    features = (None, "allocator-metrics", "capture-profile")
    for index, row in enumerate(rows):
        assert isinstance(row, dict)
        started, ended = row.get("started"), row.get("ended")
        assert isinstance(started, (int, float)) and isinstance(ended, (int, float))
        assert started <= ended and (previous is None or previous <= started)
        previous = ended
        retained_from_archive(build_root, row["log"], f"setup build {number}.{index}.log")
        command = row.get("command")
        assert isinstance(command, list) and command[:3] == ["cargo", "build", "--offline"]
        assert "--release" in command and "--locked" in command
        if features[index] is None:
            assert "--features" not in command
        else:
            assert command[command.index("--features") + 1] == features[index]
    assert build.get("probe") == {
        f"probe-src/{name}": digest for name, digest in setup_probe.items()
        if name in {"Cargo.toml.template", "src/allocation_metrics.rs",
                    "src/counting_allocator.rs", "src/main.rs"}
    }
    assert build.get("environment", {}).get("CARGO_BUILD_JOBS") == "2"
    assert build.get("environment", {}).get("CARGO_INCREMENTAL") == "0"
    binaries = build.get("binaries")
    assert isinstance(binaries, dict) and set(binaries) == {"allocation", "native", "profile"}

    retained = value.get("retained_binaries")
    assert isinstance(retained, list) and len(retained) == 3
    by_kind = {}
    for item in retained:
        assert isinstance(item, dict) and item.get("kind") in binaries
        kind = item["kind"]
        assert kind not in by_kind
        original, saved = item.get("original"), item.get("retained")
        assert isinstance(original, dict) and isinstance(saved, dict)
        assert original == binaries[kind]
        assert original.get("path") == str(c.TARGET / f"before-{kind}")
        assert saved.get("path") == str(c.TARGET / f"before-{kind}-setup-{number}")
        assert saved.get("bytes") == original.get("bytes")
        assert saved.get("sha256") == original.get("sha256")
        if Path(saved["path"]).is_file():
            external_identity(saved, f"setup retained {number}.{kind}")
        else:
            assert cleanup is not None
            assert saved in cleanup["value"].get("removed_binaries", [])
        by_kind[kind] = saved

    inputs = c.read(quality_root / "inputs.json")
    assert inputs.get("source") == c.read(archived_source)
    assert inputs.get("probe") == setup_probe
    if quality_succeeded:
        complete = c.read(quality_root / "complete.json")
        retained_artifact(quality_root, complete["inputs"],
                          f"setup quality {number}.complete.inputs")
        retained_artifact(quality_root, complete["receipts"],
                          f"setup quality {number}.complete.receipts")
    driver = inputs.get("driver")
    assert isinstance(driver, dict)
    retained_driver = P / "probe_quality.py"
    assert driver.get("bytes") == retained_driver.stat().st_size
    assert driver.get("sha256") == c.sha(retained_driver)
    qrows = c.read(quality_root / "receipts.json")
    assert isinstance(qrows, list) and qrows
    actual_first = next((index for index, row in enumerate(qrows)
                         if row.get("exit_code") != 0), None)
    if quality_succeeded:
        assert actual_first is None and len(qrows) == 3
        assert [row.get("exit_code") for row in qrows] == [0, 0, 0]
    else:
        assert first_failed == actual_first and len(qrows) == first_failed + 1
        assert [row.get("exit_code") for row in qrows[:first_failed]] == [0] * first_failed
        assert qrows[first_failed].get("exit_code") != 0
    previous = None
    test_summary = None
    for index, row in enumerate(qrows):
        assert isinstance(row, dict)
        started, ended = row.get("started"), row.get("ended")
        assert isinstance(started, (int, float)) and isinstance(ended, (int, float))
        assert started <= ended and (previous is None or previous <= started)
        previous = ended
        retained_from_archive(quality_root, row["log"], f"setup quality {number}.{index}.log")
        text = (quality_root / Path(row["log"]["path"]).name).read_text(errors="replace")
        matches = list(TEST_RESULT.finditer(text))
        if matches:
            counts = [tuple(map(int, match.groups())) for match in matches]
            test_summary = {
                "passed": sum(item[0] for item in counts),
                "failed": sum(item[1] for item in counts),
                "ignored": sum(item[2] for item in counts),
            }
    expected_tests = value.get("tests")
    assert isinstance(expected_tests, dict)
    for field in ("passed", "failed", "ignored"):
        assert isinstance(expected_tests.get(field), int) and expected_tests[field] >= 0
    if test_summary is not None:
        assert test_summary == {field: expected_tests[field]
                                for field in ("passed", "failed", "ignored")}

    qualification = None
    if quality_succeeded:
        qualification_accepted = value.get("qualification_accepted")
        assert isinstance(qualification_accepted, bool)
        assert value.get("qualification_reports") == QUALIFICATION_COUNT
        assert value.get("qualification_samples") == QUALIFICATION_COUNT
        qualification = check_setup_qualification(
            number, value, c.read(archived_source), binaries["allocation"],
            qualification_accepted
        )
    return {
        "id": path.stem,
        "attempt": artifact(path),
        "build": archive_record(build_root),
        "probe": archive_record(probe_root),
        "quality": archive_record(quality_root),
        "first_failed_gate": first_failed,
        "tests": expected_tests,
        "retained_binaries": retained,
        **({"qualification": qualification} if qualification is not None else {}),
        **({"qualification_witness": artifact(P / "qualification-schema-mismatch-0.json")}
           if qualification is not None and not qualification["accepted"] else {}),
    }


def failed_patches() -> dict[str, dict[str, Any]]:
    return {
        path.name: artifact(path)
        for path in sorted(P.glob("*-failed-*.patch"))
    }


def result() -> dict[str, Any]:
    roots = attempt_roots()
    setup_cleanup_record = setup_cleanup()
    setup = [check_setup_attempt(number, path, setup_cleanup_record)
             for number, path in setup_attempts()]
    if setup_cleanup_record is not None:
        expected = {
            (item["retained"]["path"], item["retained"]["bytes"],
             item["retained"]["sha256"])
            for attempt in setup for item in attempt["retained_binaries"]
        }
        actual = c.read(P / "setup-cleanup.json").get("removed_binaries", [])
        assert {(item.get("path"), item.get("bytes"), item.get("sha256"))
                for item in actual} == expected
        assert len(actual) == len(expected)
    main_cleanup = P / "cleanup.json"
    if main_cleanup.is_file():
        value = c.read(main_cleanup)
        assert isinstance(value, dict)
        assert set(value) == {"removed_binaries", "removed_target_bytes", "target",
                              "target_removed"}
        assert value.get("target") == str(c.TARGET)
        assert value.get("target_removed") is True and not c.TARGET.exists()
        assert isinstance(value.get("removed_target_bytes"), int)
        removed = value.get("removed_binaries")
        assert isinstance(removed, list) and len(removed) == 6
        expected_names = {f"{leg}-{kind}" for leg in ("before", "after")
                          for kind in ("native", "allocation", "profile")}
        assert {item.get("path") for item in removed if isinstance(item, dict)} == {
            str(c.TARGET / name) for name in expected_names
        }
    return {
        "schema": "litchi.performance.0806.failure-audit.v1",
        "production_adoption": False,
        "production_source_changed": False,
        "failed_attempts": [audit_attempt(root) for root in roots],
        "quality_failure": quality_failure(),
        "quality_failure_one": quality_failure_one(),
        "setup_attempts": setup,
        "setup_cleanup": setup_cleanup_record,
        "patches": failed_patches(),
        "path_policy": (
            "top-level failed-attempt roots and numbered setup archives are audited; "
            "no failed command is rerun; successful setup builds remain separate "
            "from failed quality receipts"
        ),
    }


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--write", action="store_true")
    group.add_argument("--check", action="store_true")
    args = parser.parse_args()
    output = P / "failure-audit.json"
    expected = result()
    if args.write:
        assert not output.exists(), output
        c.write(output, expected)
    else:
        assert c.read(output) == expected
    print(f"0806 failure audit PASS: {len(expected['failed_attempts'])} retained attempts")
