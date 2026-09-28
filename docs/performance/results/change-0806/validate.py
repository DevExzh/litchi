"""Fail-closed offline validator for the complete 0806 workflow packet.

The validator consumes receipts and retained reports only.  It does not build,
run, profile, or recreate a fixture.  The independent analyzers are replayed
before this reader accepts the aggregate counts or disposition.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import random
import re
import statistics
import subprocess
import sys
from pathlib import Path
from typing import Any

import custody as c


P = c.P
ROOT = c.ROOT
TARGET = c.TARGET
MAIN_BINARY_COUNT = 6
NATIVE_CHILDREN = 216
NATIVE_SAMPLES = 6_480
ALLOCATION_CHILDREN = 72
ALLOCATION_SAMPLES = 216
QUALIFICATION_CHILDREN = 18
QUALIFICATION_SAMPLES = 18
MAIN_REPORTS = 306
MAIN_SAMPLES = 6_714
PROFILE_REPORTS = 4
PROFILE_SAMPLES = 4
CROSS_REPORTS = 13
CROSS_SAMPLES = 2_888
TOTAL_REPORTS = MAIN_REPORTS + PROFILE_REPORTS + CROSS_REPORTS
TOTAL_SAMPLES = MAIN_SAMPLES + PROFILE_SAMPLES + CROSS_SAMPLES
SHAPES = ("tiny", "medium", "large", "vendor", "unicode-vendor", "valid-4attr")
MODES = ("capture", "commit", "lifecycle")
CASES = tuple((shape, mode) for shape in SHAPES for mode in MODES)
ORDERS = (
    ("before", "after"),
    ("after", "before"),
    ("before", "after"),
    ("after", "before"),
    ("after", "before"),
    ("before", "after"),
)
SOURCE_ALLOWLIST = {
    "crates/litchi-ole-common/src/xml_attributes.rs",
    "crates/litchi-opc/src/xml_attributes.rs",
    "crates/litchi-opc/src/xml_attributes/tests.rs",
    "crates/litchi-sign/src/xml_attributes.rs",
    "crates/litchi-xldm/src/xml_attributes.rs",
    "crates/xml-minifier/src/xml_attributes.rs",
}
SETUP_ATTEMPT = re.compile(r"^setup-attempt-([0-9]+)\.json$")
SETUP_PROBE_FILES = {
    "Cargo.lock",
    "Cargo.toml",
    "Cargo.toml.template",
    "src/allocation_metrics.rs",
    "src/counting_allocator.rs",
    "src/main.rs",
}
TEST_RESULT = re.compile(
    r"^test result: (?:ok|FAILED)\.\s+(\d+) passed;\s+(\d+) failed;\s+"
    r"(\d+) ignored;\s+(\d+) measured;\s+(\d+) filtered out;",
    re.MULTILINE,
)
QUALITY_FAILURE_FILES = {"00.log", "01.log", "checks.json", "source.json"}
AMENDMENT_HELPER_FILES = {
    "crates/litchi-ole-common/src/xml_attributes.rs",
    "crates/litchi-opc/src/xml_attributes.rs",
    "crates/litchi-sign/src/xml_attributes.rs",
    "crates/litchi-xldm/src/xml_attributes.rs",
    "crates/xml-minifier/src/xml_attributes.rs",
}
AMENDMENT_SHARED_TEST = "crates/litchi-opc/src/xml_attributes/tests.rs"
AMENDMENT_CHANGES = {
    "litchi-ole-common-xml_attributes.rs": "crates/litchi-ole-common/src/xml_attributes.rs",
    "litchi-opc-xml_attributes.rs": "crates/litchi-opc/src/xml_attributes.rs",
    "litchi-sign-xml_attributes.rs": "crates/litchi-sign/src/xml_attributes.rs",
    "litchi-xldm-xml_attributes.rs": "crates/litchi-xldm/src/xml_attributes.rs",
    "xml-minifier-xml_attributes.rs": "crates/xml-minifier/src/xml_attributes.rs",
}
CANDIDATE_CHANGES = {
    **AMENDMENT_CHANGES,
    "litchi-opc-xml_attributes-tests.rs": AMENDMENT_SHARED_TEST,
}
VISIBILITY_ARCHIVE = "litchi-ole-common-xml_attributes.rs"
VISIBILITY_PRODUCTION = "crates/litchi-ole-common/src/xml_attributes.rs"
VISIBILITY_FILES = {VISIBILITY_PRODUCTION}


def fail(message: str) -> None:
    raise AssertionError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def read(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing evidence: {path}")
    try:
        return c.read(path)
    except (OSError, ValueError) as error:
        fail(f"invalid JSON {path}: {error}")


def packet_path(raw: str | Path) -> Path:
    path = Path(raw)
    if path.is_absolute():
        if path.is_file():
            return path.resolve()
        marker = "/docs/performance/results/change-0806/"
        text = path.as_posix()
        if marker in text:
            return (P / text.split(marker, 1)[1]).resolve()
        return path.resolve()
    for candidate in (P / path, ROOT / path):
        if candidate.exists():
            return candidate.resolve()
    return (P / path).resolve()


def descriptor(value: Any, label: str, *, allow_missing: bool = False,
               external: bool = False) -> Path | None:
    require(isinstance(value, dict), f"{label} is not an artifact descriptor")
    raw = value.get("path")
    require(isinstance(raw, str) and raw, f"{label}.path is missing")
    size, digest = value.get("bytes"), value.get("sha256")
    require(isinstance(size, int) and not isinstance(size, bool) and size >= 0,
            f"{label}.bytes is invalid")
    require(isinstance(digest, str) and len(digest) == 64,
            f"{label}.sha256 is invalid")
    path = packet_path(raw)
    if not path.is_file() or path.is_symlink():
        if allow_missing:
            return None
        fail(f"missing {label}: {raw}")
    actual = c.artifact(path)
    require(actual["bytes"] == size and actual["sha256"] == digest,
            f"{label} identity changed")
    if not external:
        require(path.is_relative_to(P), f"{label} escaped packet")
    return path


def source_manifest(path: Path, label: str) -> dict[str, Any]:
    value = read(path)
    require(isinstance(value, dict), f"{label} is malformed")
    revision, files = value.get("revision"), value.get("files")
    require(isinstance(revision, str) and len(revision) == 40,
            f"{label}.revision is invalid")
    require(isinstance(files, dict) and files, f"{label}.files is missing")
    require(all(isinstance(name, str) and isinstance(digest, str)
                and len(digest) == 64 for name, digest in files.items()),
            f"{label}.files contains an invalid digest")
    return {"revision": revision, "files": dict(files)}


def current_source() -> dict[str, str]:
    return c.source()["files"]


def source_descriptor(value: Any, label: str) -> tuple[Path, dict[str, Any]]:
    path = descriptor(value, label)
    require(path is not None, label)
    return path, source_manifest(path, label)


def retained_log(root: Path, value: Any, label: str) -> Path:
    """Resolve a receipt log retained under an archive directory.

    Setup receipts deliberately retain their original absolute path.  The
    archived copy is keyed by that path's basename, so validation must bind
    the copied bytes instead of silently following a later live log.
    """

    require(isinstance(value, dict), f"{label} is malformed")
    raw = value.get("path")
    require(isinstance(raw, str) and Path(raw).is_absolute(),
            f"{label}.path is not absolute")
    path = root / Path(raw).name
    require(path.is_file() and not path.is_symlink(), f"missing {label}: {path}")
    actual = c.artifact(path)
    require(actual["bytes"] == value.get("bytes")
            and actual["sha256"] == value.get("sha256"),
            f"{label} identity changed")
    return path


def retained_artifact(root: Path, value: Any, label: str) -> Path:
    """Bind a copied archive artifact to the receipt's original descriptor."""

    require(isinstance(value, dict), f"{label} is malformed")
    raw = value.get("path")
    require(isinstance(raw, str) and Path(raw).is_absolute(),
            f"{label}.path is not absolute")
    path = root / Path(raw).name
    require(path.is_file() and not path.is_symlink(), f"missing {label}: {path}")
    actual = c.artifact(path)
    require(actual["bytes"] == value.get("bytes")
            and actual["sha256"] == value.get("sha256"),
            f"{label} identity changed")
    return path


def setup_attempt_paths() -> list[tuple[int, Path]]:
    found = []
    for path in sorted(P.iterdir()):
        match = SETUP_ATTEMPT.fullmatch(path.name)
        if match is not None and path.is_file() and not path.is_symlink():
            found.append((int(match.group(1)), path))
    require([number for number, _ in found] == list(range(len(found))),
            "setup attempt numbering has a gap")
    require(found, "setup attempt archives are missing")
    return found


def check_setup_cleanup(expected: list[dict[str, Any]], require_final: bool) -> dict[str, Any] | None:
    """Check the setup binaries separately from the six main binaries."""

    path = P / "setup-cleanup.json"
    if not path.is_file():
        require(not require_final, "setup cleanup witness is missing")
        require(TARGET.exists(), "setup binaries disappeared without setup cleanup")
        for item in expected:
            retained = Path(item["retained"]["path"])
            require(retained.is_file() and c.artifact(retained) == item["retained"],
                    f"setup retained binary changed: {retained}")
        return None
    value = read(path)
    require(value.get("target") == str(TARGET)
            and value.get("target_removed") is True
            and not TARGET.exists(), "setup cleanup is incomplete")
    if "schema" in value:
        require(value["schema"] == "litchi.performance.0806.setup-cleanup.v1",
                "setup cleanup schema changed")
    removed = value.get("removed_binaries")
    require(isinstance(removed, list), "setup cleanup binary list is missing")
    expected_ids = {(item["retained"]["path"], item["retained"]["bytes"],
                     item["retained"]["sha256"]) for item in expected}
    actual_ids = {(item.get("path"), item.get("bytes"), item.get("sha256"))
                  for item in removed if isinstance(item, dict)}
    require(actual_ids == expected_ids and len(removed) == len(expected),
            "setup cleanup identities changed")
    for item in removed:
        require(not Path(item["path"]).exists(),
                f"setup binary remains after cleanup: {item['path']}")
    return {"path": str(path.relative_to(P)), "value": value}


def check_setup_qualification(number: int, attempt: dict[str, Any],
                              before: dict[str, Any], allocation_binary: dict[str, Any],
                              accepted: bool) -> dict[str, Any]:
    """Audit the successful probe run whose qualification metadata was rejected."""

    root = P / attempt["qualification_archive"]
    require(root.is_dir() and not root.is_symlink(),
            f"setup qualification {number} archive is missing")
    complete = read(root / "complete.json")
    require(complete.get("children") == QUALIFICATION_CHILDREN,
            f"setup qualification {number} child count changed")
    source_path = retained_artifact(root, complete.get("source"),
                                    f"setup qualification {number}.source")
    require(source_manifest(source_path, f"setup qualification {number}.source") == before,
            f"setup qualification {number} source changed")
    retained_artifact(root, complete.get("receipts"),
                      f"setup qualification {number}.receipts")
    rows = read(root / "receipts.json")
    require(isinstance(rows, list) and len(rows) == QUALIFICATION_CHILDREN,
            f"setup qualification {number} receipt count changed")
    expected = [(0, shape, mode, "before") for shape, mode in CASES]
    mismatch_reports = []
    for index, (row, identity) in enumerate(zip(rows, expected)):
        block, shape, mode, leg = identity
        require(isinstance(row, dict)
                and (row.get("lane"), row.get("block"), row.get("shape"),
                     row.get("mode"), row.get("leg"))
                == ("qualification", block, shape, mode, leg)
                and row.get("exit_code") == 0,
                f"setup qualification {number} row {index} changed")
        require(row.get("binary") == allocation_binary,
                f"setup qualification {number} row {index} binary changed")
        require(row.get("started") <= row.get("ended"),
                f"setup qualification {number} row {index} timing changed")
        log = retained_artifact(root, row.get("log"),
                                f"setup qualification {number} row {index}.log")
        report_path = retained_artifact(root, row.get("report"),
                                       f"setup qualification {number} row {index}.report")
        rss = retained_artifact(root, row.get("rss"),
                                f"setup qualification {number} row {index}.rss")
        require(rss.read_text().strip().isdigit(),
                f"setup qualification {number} row {index} RSS changed")
        command = row.get("command")
        require(isinstance(command, list)
                and command[0:4] == ["/usr/bin/time", "-f", "%M", "-o"]
                and "taskset" in command and "-c" in command
                and command[command.index("-c") + 1] == "12"
                and "--mode" in command
                and command[command.index("--mode") + 1] == mode
                and "--shape" in command
                and command[command.index("--shape") + 1] == shape
                and "--samples" in command
                and command[command.index("--samples") + 1] == "1"
                and "--warmup" in command
                and command[command.index("--warmup") + 1] == "0",
                f"setup qualification {number} row {index} command changed")
        require("--output" in command
                and "-o" in command
                and Path(command[command.index("--output") + 1]).resolve()
                == Path(row["report"]["path"]).resolve()
                and command[command.index("-o") + 1] == row["rss"]["path"]
                and allocation_binary["path"] in command,
                f"setup qualification {number} row {index} artifact binding changed")
        raw = read(report_path)
        actual_shape = shape if accepted or shape != "valid-4attr" else "valid4-attr"
        require(raw.get("schema") == "litchi.pptx.public-workflow-probe-0806.v1"
                and raw.get("tool") == "public-pptx-probe-0806"
                and raw.get("mode") == mode and raw.get("shape") == actual_shape
                and raw.get("samples_requested") == 1
                and isinstance(raw.get("samples"), list)
                and len(raw["samples"]) == 1,
                f"setup qualification {number} row {index} report metadata changed")
        sample = raw["samples"][0]
        verification = sample.get("verification")
        require(isinstance(verification, dict)
                and verification.get("semantic_check") is True
                and verification.get("reopened") is True
                and verification.get("expected_text") == verification.get("actual_text"),
                f"setup qualification {number} row {index} preservation failed")
        if actual_shape != shape:
            mismatch_reports.append(report_path.name)
    if accepted:
        require(not mismatch_reports,
                "accepted setup qualification still has metadata mismatches")
    else:
        mismatch = read(P / "qualification-schema-mismatch-0.json")
        require(mismatch.get("schema")
                == "litchi.performance.0806.qualification-schema-mismatch.v1"
                and mismatch.get("passed") is False
                and mismatch.get("capture_children_succeeded") == QUALIFICATION_CHILDREN
                and isinstance(mismatch.get("detected_by"), str)
                and isinstance(mismatch.get("explanation"), str)
                and mismatch["detected_by"].strip()
                and mismatch["explanation"].strip()
                and len(mismatch.get("metadata_mismatches", [])) == 3,
                "qualification schema mismatch witness changed")
        witness_reports = []
        for index, item in enumerate(mismatch["metadata_mismatches"]):
            require(item.get("actual_serialized_shape") == "valid4-attr"
                    and item.get("expected_cli_and_plan_shape") == "valid-4attr",
                    f"qualification mismatch {index} changed")
            report = retained_artifact(root, item.get("report"),
                                       f"qualification mismatch {index}.report")
            witness_reports.append(report.name)
        require(sorted(witness_reports) == sorted(mismatch_reports)
                and sorted(mismatch_reports) == sorted([
                    f"0-valid-4attr-{mode}-before.json" for mode in MODES
                ]), "qualification mismatch reports changed")
    return {"reports": QUALIFICATION_CHILDREN, "samples": QUALIFICATION_SAMPLES,
            "accepted": accepted, "metadata_mismatches": len(mismatch_reports)}


def check_setup_archives(before: dict[str, Any], require_final: bool) -> dict[str, Any]:
    attempts = []
    expected_retained: list[dict[str, Any]] = []
    previous_archived = None
    for number, path in setup_attempt_paths():
        value = read(path)
        require(value.get("schema") == "litchi.performance.0806.setup-attempt.v1",
                f"setup attempt {number} schema changed")
        quality_succeeded = value.get("probe_quality_succeeded") is True
        require(value.get("build_succeeded") is True
                and value.get("probe_quality_succeeded") is quality_succeeded
                and value.get("cross_lane_affected") is False,
                f"setup attempt {number} status changed")
        started_fields = [value[name] for name in
                          ("main_captures_started", "main_timed_captures_started")
                          if name in value]
        require(started_fields and all(flag is False for flag in started_fields),
                f"setup attempt {number} capture status changed")
        archived_at = value.get("archived_at")
        require(isinstance(archived_at, (int, float))
                and (previous_archived is None or previous_archived <= archived_at),
                f"setup attempt {number} chronology changed")
        previous_archived = archived_at
        require(isinstance(value.get("reason"), str) and value["reason"].strip(),
                f"setup attempt {number} has no actual failure reason")
        expected_quality_archive = (f"probe-quality-before-setup-{number}"
                                    if quality_succeeded else
                                    f"probe-quality-before-failed-{number}")
        require(value.get("build_archive") == f"build-before-setup-{number}"
                and value.get("probe_archive") == f"probe-src-setup-{number}"
                and value.get("quality_archive") == expected_quality_archive,
                f"setup attempt {number} archive mapping changed")
        first_failed = value.get("first_failed_gate")
        if quality_succeeded:
            require(first_failed is None and "first_failed_gate" not in value,
                    f"setup attempt {number} successful quality status changed")
        else:
            require(isinstance(first_failed, int) and first_failed >= 0,
                    f"setup attempt {number} failed gate is invalid")
        production = value.get("production_source")
        source_path = descriptor(production, f"setup attempt {number}.production_source")
        require(source_path is not None and source_manifest(source_path,
                                                            f"setup attempt {number}.source") == before,
                f"setup attempt {number} source changed")

        build_root = P / value["build_archive"]
        probe_root = P / value["probe_archive"]
        quality_root = P / value["quality_archive"]
        for root, label in ((build_root, "build"), (probe_root, "probe"),
                            (quality_root, "quality")):
            require(root.is_dir() and not root.is_symlink(),
                    f"setup attempt {number} {label} archive is missing")

        probe_files = {str(item.relative_to(probe_root)): c.sha(item)
                       for item in probe_root.rglob("*")
                       if item.is_file() and not item.is_symlink()}
        require(set(probe_files) == SETUP_PROBE_FILES,
                f"setup attempt {number} probe source inventory changed")
        probe = c.read(build_root / "probe.json")
        require(probe == {f"probe-src/{name}": digest for name, digest in probe_files.items()
                          if name in {"Cargo.toml.template", "src/allocation_metrics.rs",
                                      "src/counting_allocator.rs", "src/main.rs"}},
                f"setup attempt {number} build probe custody changed")

        build = read(build_root / "build.json")
        build_source = descriptor(build.get("source"),
                                  f"setup build {number}.source")
        require(build_source is not None
                and source_manifest(build_source, f"setup build {number}.source") == before,
                f"setup build {number} source changed")
        archived_source = source_manifest(build_root / "source.json",
                                          f"setup build {number}.archived_source")
        require(archived_source == before, f"setup build {number} archived source changed")
        rows = read(build_root / "commands.json")
        require(isinstance(rows, list) and len(rows) == 3,
                f"setup build {number} command count changed")
        require(build.get("environment", {}).get("CARGO_BUILD_JOBS") == "2"
                and build.get("environment", {}).get("CARGO_INCREMENTAL") == "0",
                f"setup build {number} environment changed")
        previous = None
        expected_features = (None, "allocator-metrics", "capture-profile")
        for index, row in enumerate(rows):
            require(isinstance(row, dict) and row.get("exit_code") == 0,
                    f"setup build {number}.{index} failed")
            require(isinstance(row.get("started"), (int, float))
                    and isinstance(row.get("ended"), (int, float))
                    and row["started"] <= row["ended"]
                    and (previous is None or previous <= row["started"]),
                    f"setup build {number}.{index} timing changed")
            previous = row["ended"]
            retained_log(build_root, row.get("log"),
                         f"setup build {number}.{index}.log")
            command = row.get("command")
            require(isinstance(command, list)
                    and command[:3] == ["cargo", "build", "--offline"]
                    and "--release" in command and "--locked" in command,
                    f"setup build {number}.{index} command changed")
            if expected_features[index] is None:
                require("--features" not in command,
                        f"setup build {number}.{index} feature changed")
            else:
                require("--features" in command
                        and command[command.index("--features") + 1]
                        == expected_features[index],
                        f"setup build {number}.{index} feature changed")
        binaries = build.get("binaries")
        require(isinstance(binaries, dict)
                and set(binaries) == {"allocation", "native", "profile"},
                f"setup build {number} binary set changed")
        retained = value.get("retained_binaries")
        require(isinstance(retained, list) and len(retained) == 3,
                f"setup attempt {number} binary custody changed")
        kinds = set()
        for item in retained:
            require(isinstance(item, dict) and item.get("kind") in binaries,
                    f"setup attempt {number} binary record malformed")
            kind = item["kind"]
            require(kind not in kinds, f"setup attempt {number} duplicate binary")
            kinds.add(kind)
            original, saved = item.get("original"), item.get("retained")
            require(isinstance(original, dict) and isinstance(saved, dict),
                    f"setup attempt {number}.{kind} binary record malformed")
            require(original == binaries[kind]
                    and original.get("path") == str(TARGET / f"before-{kind}")
                    and saved.get("path") == str(TARGET / f"before-{kind}-setup-{number}")
                    and saved.get("bytes") == original.get("bytes")
                    and saved.get("sha256") == original.get("sha256"),
                    f"setup attempt {number}.{kind} binary identity changed")
            expected_retained.append({"kind": kind, "retained": saved})
        frozen = read(build_root / "frozen-inputs.json")
        require(isinstance(frozen, dict), f"setup build {number} frozen inputs missing")
        for name, digest in frozen.items():
            require(isinstance(digest, str) and len(digest) == 64
                    and (P / name).is_file() and c.sha(P / name) == digest,
                    f"setup build {number} frozen input changed: {name}")

        inputs = read(quality_root / "inputs.json")
        require(inputs.get("source") == c.read(P / "source.json")
                and inputs.get("probe") == probe_files,
                f"setup quality {number} inputs changed")
        if quality_succeeded:
            complete = read(quality_root / "complete.json")
            retained_artifact(quality_root, complete.get("inputs"),
                              f"setup quality {number}.complete.inputs")
            retained_artifact(quality_root, complete.get("receipts"),
                              f"setup quality {number}.complete.receipts")
        driver = descriptor(inputs.get("driver"), f"setup quality {number}.driver")
        require(driver is not None and c.sha(driver) == c.sha(P / "probe_quality.py"),
                f"setup quality {number} driver changed")
        qrows = read(quality_root / "receipts.json")
        require(isinstance(qrows, list) and qrows,
                f"setup quality {number} receipts missing")
        actual_first = next((index for index, row in enumerate(qrows)
                             if row.get("exit_code") != 0), None)
        if quality_succeeded:
            require(actual_first is None and len(qrows) == 3
                    and [row.get("exit_code") for row in qrows] == [0, 0, 0],
                    f"setup quality {number} successful receipts changed")
        else:
            require(actual_first == first_failed and len(qrows) == first_failed + 1,
                    f"setup quality {number} failed gate does not match receipt")
            require([row.get("exit_code") for row in qrows[:first_failed]]
                    == [0] * first_failed,
                    f"setup quality {number} earlier gate failed")
        previous = None
        test_summary = None
        for index, row in enumerate(qrows):
            require(isinstance(row, dict)
                    and isinstance(row.get("started"), (int, float))
                    and isinstance(row.get("ended"), (int, float))
                    and row["started"] <= row["ended"]
                    and (previous is None or previous <= row["started"]),
                    f"setup quality {number}.{index} timing changed")
            previous = row["ended"]
            log = retained_log(quality_root, row.get("log"),
                               f"setup quality {number}.{index}.log")
            matches = list(TEST_RESULT.finditer(log.read_text(errors="replace")))
            if matches:
                counts = [tuple(map(int, match.groups())) for match in matches]
                test_summary = {
                    "passed": sum(item[0] for item in counts),
                    "failed": sum(item[1] for item in counts),
                    "ignored": sum(item[2] for item in counts),
                }
        tests = value.get("tests")
        require(isinstance(tests, dict), f"setup attempt {number} tests missing")
        for field in ("passed", "failed", "ignored"):
            require(isinstance(tests.get(field), int) and tests[field] >= 0,
                    f"setup attempt {number} test count malformed")
        if test_summary is not None:
            require(test_summary == {field: tests[field]
                                     for field in ("passed", "failed", "ignored")},
                    f"setup attempt {number} test receipt changed")
        if quality_succeeded:
            qualification_accepted = value.get("qualification_accepted")
            require(isinstance(qualification_accepted, bool)
                    and value.get("qualification_reports") == QUALIFICATION_CHILDREN
                    and value.get("qualification_samples") == QUALIFICATION_SAMPLES,
                    f"setup attempt {number} qualification status changed")
            qualification = check_setup_qualification(number, value, before,
                                                       binaries["allocation"],
                                                       qualification_accepted)
        else:
            qualification = None
        attempts.append({"id": path.stem, "first_failed_gate": first_failed,
                         "tests": tests, "retained_binaries": retained})
        if qualification is not None:
            attempts[-1]["qualification"] = qualification

    cleanup = check_setup_cleanup(expected_retained, require_final)
    return {"attempts": attempts, "cleanup": cleanup,
            "retained_binaries": len(expected_retained)}


def run_reader(script: str) -> None:
    result = subprocess.run([sys.executable, "-B", str(P / script), "--check"],
                            cwd=ROOT, stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE, text=True)
    require(result.returncode == 0,
            f"{script} replay failed: {result.stderr.strip() or result.stdout.strip()}")


def check_workspace() -> dict[str, Any]:
    origin = read(P / "origin.json")
    require(origin.get("base") == "e3ff267ee3454e71d66f177f54f3cd05e0d9cce5",
            "origin base changed")
    require(set(origin) == {"base", "unrelated", "worktrees"},
            "origin schema changed")
    require(TARGET.is_absolute() and TARGET.name == "litchi-target-0806",
            "custody target convention changed")
    subprocess.run(["git", "merge-base", "--is-ancestor", origin["base"], "HEAD"],
                   cwd=ROOT, check=True)
    for name, digest in origin.get("unrelated", {}).items():
        path = ROOT / name
        require(path.is_file() and c.sha(path) == digest,
                f"unrelated workspace file changed: {name}")
    recorded = origin.get("worktrees", "").strip().split("\n\n")
    actual = subprocess.check_output(["git", "worktree", "list", "--porcelain"],
                                     cwd=ROOT, text=True).strip().split("\n\n")

    def blocks(values: list[str]) -> dict[str, str]:
        return {block.splitlines()[0][9:]: block for block in values if block}

    old, new = blocks(recorded), blocks(actual)
    require(ROOT.as_posix() in new, "main worktree disappeared")
    for path, block in old.items():
        if Path(path).resolve() != ROOT.resolve():
            require(new.get(path) == block, f"unrelated worktree changed: {path}")
    lock = P / "workspace-Cargo.lock"
    require(lock.is_file() and c.sha(lock) == c.sha(ROOT / "Cargo.lock"),
            "workspace Cargo.lock changed")
    return {"unrelated_files": len(origin.get("unrelated", {})),
            "other_worktrees": max(0, len(old) - 1)}


def check_candidate(before: dict[str, Any], after: dict[str, Any]) -> dict[str, Any]:
    changed = {name for name in set(before["files"]) | set(after["files"])
               if before["files"].get(name) != after["files"].get(name)}
    require(changed == SOURCE_ALLOWLIST,
            f"candidate source change set changed: {sorted(changed)}")
    require("crates/litchi-formula/src/omml/xml_attributes.rs" not in changed,
            "formula source changed unexpectedly")
    manifest_path = P / "candidate/manifest.json"
    manifest = read(manifest_path)
    require(manifest.get("schema") == "litchi.performance.0806.workflow-candidate.v1",
            "candidate manifest schema changed")
    require(manifest.get("base_commit") == read(P / "origin.json")["base"],
            "candidate base changed")
    files = manifest.get("files")
    require(isinstance(files, dict) and len(files) == len(SOURCE_ALLOWLIST),
            "candidate manifest files are malformed")
    require(set(files) == set(CANDIDATE_CHANGES),
            "candidate manifest archive names changed")
    paths = {entry.get("production_path") for entry in files.values()
             if isinstance(entry, dict)}
    require(paths == changed, "candidate manifest paths differ from build source")
    for entry in files.values():
        require(isinstance(entry, dict), "candidate manifest entry malformed")
        require(set(entry) == {"production_path", "before", "after"},
                "candidate manifest entry fields changed")
        for leg, expected in (("before", before), ("after", after)):
            item = entry.get(leg)
            path = descriptor(item, f"candidate.{leg}")
            require(path is not None, f"candidate.{leg} missing")
            production = entry["production_path"]
            require(c.sha(path) == expected["files"][production],
                    f"candidate archive {leg} differs: {production}")
    patch = P / "candidate/candidate.patch"
    design = P / "candidate/design.md"
    require(patch.is_file() and design.is_file(), "candidate patch/design missing")
    application_path = P / "application.json"
    application = read(application_path)
    app_source = application.get("source")
    require(isinstance(app_source, dict)
            and app_source.get("revision") == before["revision"],
            "application source revision differs from baseline")
    candidate_files = app_source.get("files")
    require(isinstance(candidate_files, dict)
            and set(candidate_files) == set(before["files"]),
            "application source file census changed")
    candidate_changed = {
        name for name in set(before["files"]) | set(candidate_files)
        if before["files"].get(name) != candidate_files.get(name)
    }
    require(candidate_changed == SOURCE_ALLOWLIST,
            "application source change set differs from candidate")
    require(all(candidate_files[name] == before["files"][name]
               for name in set(before["files"]) - SOURCE_ALLOWLIST),
            "application source changed outside candidate scope")
    require(any(isinstance(entry, dict)
                and entry.get("production_path") in SOURCE_ALLOWLIST
                and entry.get("after", {}).get("sha256") == candidate_files.get(
                    entry.get("production_path"))
                for entry in files.values()),
            "application source is not bound to candidate archives")
    app_manifest = descriptor(application.get("manifest"), "application.manifest")
    app_patch = descriptor(application.get("patch"), "application.patch")
    require(app_manifest is not None and app_patch is not None,
            "application witnesses are incomplete")
    return {"changed_files": sorted(changed),
            "manifest": str(manifest_path.relative_to(P)),
            "application": str(application_path.relative_to(P)),
            "source": "application.source",
            "candidate_source": app_source}


def cleanup_records() -> tuple[dict[str, Any] | None, bool]:
    path = P / "cleanup.json"
    if not path.is_file():
        return None, False
    value = read(path)
    require(isinstance(value, dict), "cleanup witness is malformed")
    require(set(value) == {"removed_binaries", "removed_target_bytes", "target",
                            "target_removed"}
            and value.get("target") == str(TARGET)
            and value.get("target_removed") is True
            and not TARGET.exists(), "main target cleanup is incomplete")
    require(isinstance(value.get("removed_target_bytes"), int)
            and value["removed_target_bytes"] >= 0,
            "main cleanup byte total is malformed")
    removed = value.get("removed_binaries")
    require(isinstance(removed, list) and len(removed) == MAIN_BINARY_COUNT,
            "main cleanup binary count changed")
    identities = {(item.get("path"), item.get("bytes"), item.get("sha256"))
                  for item in removed if isinstance(item, dict)}
    require(len(identities) == MAIN_BINARY_COUNT,
            "main cleanup contains duplicate binary identities")
    expected_names = {f"{leg}-{kind}" for leg in ("before", "after")
                      for kind in ("native", "allocation", "profile")}
    for item in removed:
        raw_path = item.get("path") if isinstance(item, dict) else None
        require(isinstance(item, dict)
                and isinstance(raw_path, str)
                and Path(raw_path).parent == TARGET
                and Path(raw_path).name in expected_names
                and isinstance(item.get("bytes"), int)
                and isinstance(item.get("sha256"), str)
                and len(item["sha256"]) == 64,
                "main cleanup binary identity is malformed or escaped target")
    require({Path(item["path"]).name for item in removed} == expected_names,
            "main cleanup binary names changed")
    return value, True


def check_builds(before: dict[str, Any], after: dict[str, Any], cleanup: Any,
                 cleanup_ok: bool) -> dict[str, Any]:
    builds = {}
    expected_kinds = {"native", "allocation", "profile"}
    for leg in ("before", "after"):
        root = P / f"build-{leg}"
        value = read(root / "build.json")
        path, source = source_descriptor(value.get("source"), f"build-{leg}.source")
        expected_source = before if leg == "before" else after
        require(source == expected_source, f"build-{leg} source changed")
        require(isinstance(value.get("binaries"), dict)
                and set(value["binaries"]) == expected_kinds,
                f"build-{leg} binary set changed")
        for kind, binary in value["binaries"].items():
            require(isinstance(binary, dict)
                    and binary.get("path") == str(TARGET / f"{leg}-{kind}"),
                    f"build-{leg}.{kind} binary path changed")
            descriptor(binary, f"build-{leg}.{kind}.binary", allow_missing=cleanup_ok,
                       external=True)
            if cleanup_ok:
                require(any(binary == item for item in cleanup["removed_binaries"]),
                        f"build-{leg}.{kind} missing cleanup identity")
        rows = value.get("rows")
        require(isinstance(rows, list) and len(rows) == 3,
                f"build-{leg} command count changed")
        seen = set()
        for row in rows:
            require(isinstance(row, dict) and row.get("exit_code") == 0,
                    f"build-{leg} command failed")
            command = row.get("command")
            require(isinstance(command, list) and command and command[0] == "cargo",
                    f"build-{leg} command malformed")
            if "allocator-metrics" in command:
                kind = "allocation"
            elif "capture-profile" in command:
                kind = "profile"
            else:
                kind = "native"
            require(kind not in seen, f"build-{leg} duplicate {kind} command")
            seen.add(kind)
            descriptor(row.get("log"), f"build-{leg}.{kind}.log")
            require(row.get("started") <= row.get("ended"),
                    f"build-{leg}.{kind} timing receipt malformed")
        require(seen == expected_kinds, f"build-{leg} command set changed")
        require(value.get("environment", {}).get("CARGO_BUILD_JOBS") == "2"
                and value.get("environment", {}).get("CARGO_INCREMENTAL") == "0",
                f"build-{leg} environment changed")
        frozen = read(root / "frozen-inputs.json")
        require(isinstance(frozen, dict), f"build-{leg} frozen inputs missing")
        for name, digest in frozen.items():
            require(isinstance(digest, str) and len(digest) == 64,
                    f"build-{leg} frozen digest invalid: {name}")
            require((P / name).is_file() and c.sha(P / name) == digest,
                    f"build-{leg} frozen input changed: {name}")
        builds[leg] = value
        require(path is not None, f"build-{leg} source missing")
    require(before["revision"] == read(P / "origin.json")["base"],
            "before build is not bound to origin")
    require(read(P / "build-before/frozen-inputs.json")
            == read(P / "build-after/frozen-inputs.json"),
            "build frozen inputs differ")
    return builds


def check_quality(after: dict[str, Any]) -> dict[str, Any]:
    value = read(P / "quality.json")
    path, source = source_descriptor(value.get("source"), "quality.source")
    require(source == after, "quality source differs from after build")
    rows = value.get("rows")
    require(isinstance(rows, list) and len(rows) == 6, "quality gate count changed")
    try:
        import analyze
        expected = analyze.QUALITY_COMMANDS
    except (AttributeError, ImportError):
        expected = None
    test_summary = None
    previous = None
    for index, row in enumerate(rows):
        require(isinstance(row, dict) and row.get("exit_code") == 0,
                f"quality gate {index} failed")
        if expected is not None:
            require(row.get("command") == expected[index],
                    f"quality gate {index} command changed")
        require(row.get("started") <= row.get("ended"),
                f"quality gate {index} timing malformed")
        if previous is not None:
            require(previous <= row["started"], "quality gates were not serial")
        previous = row["ended"]
        log = descriptor(row.get("log"), f"quality gate {index}.log")
        require(log is not None, f"quality gate {index} log missing")
        if index == 2:
            matches = re.findall(
                r"^test result: (?:ok|FAILED)\.\s+(\d+) passed; (\d+) failed; "
                r"(\d+) ignored; (\d+) measured; (\d+) filtered out;",
                log.read_text(errors="replace"), re.MULTILINE)
            require(matches, "quality test gate has no test-result records")
            test_summary = {
                "suites": len(matches),
                "passed": sum(int(row[0]) for row in matches),
                "failed": sum(int(row[1]) for row in matches),
                "ignored": sum(int(row[2]) for row in matches),
            }
            require(test_summary["failed"] == 0,
                    "quality test gate contains a failed suite")
    require(test_summary is not None, "quality test summary missing")
    env = value.get("environment", {})
    require(env.get("CARGO_BUILD_JOBS") == "2"
            and env.get("CARGO_INCREMENTAL") == "0"
            and env.get("RUSTDOCFLAGS") == "-D warnings",
            "quality environment changed")
    packages = [item for item in rows[0]["command"] if item == "-p"]
    require(len(packages) == 14, "quality package scope changed")
    return {"gates": len(rows), "test_summary": test_summary,
            "source": str(path.relative_to(P))}


def check_quality_failure(candidate_source: dict[str, Any]) -> dict[str, Any]:
    """Bind the original candidate's retained production-quality failure.

    The first quality run stopped at cargo check after formatting succeeded.
    It is historical evidence for the later source amendment, so this reader
    checks its source, receipt, log and actual diagnostic without treating the
    failed run as a final quality result.
    """
    root = P / "quality-0"
    require(root.is_dir() and not root.is_symlink(),
            "original quality failure archive is missing")
    files = {str(path.relative_to(root)) for path in root.rglob("*")
             if path.is_file() and not path.is_symlink()}
    require(files == QUALITY_FAILURE_FILES,
            "original quality failure archive inventory changed")
    archived_source = source_manifest(root / "source.json",
                                      "quality-0.source")
    require(archived_source == candidate_source,
            "quality-0 source is not the immutable original candidate")
    rows = read(root / "checks.json")
    require(isinstance(rows, list) and len(rows) == 2,
            "quality-0 receipt count changed")
    try:
        import analyze
        expected = analyze.QUALITY_COMMANDS[:2]
    except (AttributeError, ImportError):
        expected = None
    previous = None
    logs = []
    for index, row in enumerate(rows):
        require(isinstance(row, dict)
                and row.get("exit_code") == (0 if index == 0 else 101)
                and isinstance(row.get("started"), (int, float))
                and isinstance(row.get("ended"), (int, float))
                and row["started"] <= row["ended"]
                and (previous is None or previous <= row["started"]),
                f"quality-0 receipt {index} changed")
        previous = row["ended"]
        if expected is not None:
            require(row.get("command") == expected[index],
                    f"quality-0 command {index} changed")
        log = descriptor(row.get("log"), f"quality-0.{index}.log")
        require(log is not None and log.resolve().parent == root.resolve(),
                f"quality-0 log {index} escaped archive")
        logs.append(log)
    require(not logs[0].read_text(errors="replace").strip(),
            "quality-0 formatting failure log changed")
    failed_log = logs[1].read_text(errors="replace")
    require("method `unchecked_attributes` is never used" in failed_log
            and "crates/litchi-sign/src/xml_attributes.rs" in failed_log
            and "could not compile `litchi-sign`" in failed_log,
            "quality-0 failed diagnostic changed")
    return {"archive": "quality-0", "failed_gate": 1,
            "exit_codes": [row["exit_code"] for row in rows],
            "source": {"revision": archived_source["revision"],
                        "files": len(archived_source["files"])},
            "logs": [str(path.relative_to(P)) for path in logs]}


def check_quality_failure_one(amended_source: dict[str, Any]) -> dict[str, Any]:
    """Bind the retained full-workspace visibility regression.

    ``quality-1`` is the production-quality attempt after the five-helper
    constructor amendment.  Its source is the intermediate amendment source,
    and its nonzero all-features check is historical evidence for the separate
    public-visibility repair.  The failed receipt is retained as evidence and
    never counted as a successful quality gate.
    """
    root = P / "quality-1"
    require(root.is_dir() and not root.is_symlink(),
            "quality-1 failure archive is missing")
    files = {str(path.relative_to(root)) for path in root.rglob("*")
             if path.is_file() and not path.is_symlink()}
    require(files == QUALITY_FAILURE_FILES,
            "quality-1 failure archive inventory changed")
    archived_source = source_manifest(root / "source.json",
                                      "quality-1.source")
    require(archived_source == amended_source,
            "quality-1 source is not the constructor-amended source")
    rows = read(root / "checks.json")
    require(isinstance(rows, list) and len(rows) == 2,
            "quality-1 receipt count changed")
    try:
        import analyze
        expected = analyze.QUALITY_COMMANDS[:2]
    except (AttributeError, ImportError):
        expected = None
    previous = None
    logs = []
    for index, row in enumerate(rows):
        require(isinstance(row, dict)
                and row.get("exit_code") == (0 if index == 0 else 101)
                and isinstance(row.get("started"), (int, float))
                and isinstance(row.get("ended"), (int, float))
                and row["started"] <= row["ended"]
                and (previous is None or previous <= row["started"]),
                f"quality-1 receipt {index} changed")
        previous = row["ended"]
        if expected is not None:
            require(row.get("command") == expected[index],
                    f"quality-1 command {index} changed")
        log = descriptor(row.get("log"), f"quality-1.{index}.log")
        require(log is not None and log.resolve().parent == root.resolve(),
                f"quality-1 log {index} escaped archive")
        logs.append(log)
    require(not logs[0].read_text(errors="replace").strip(),
            "quality-1 formatting failure log changed")
    failed_log = logs[1].read_text(errors="replace")
    require("error[E0603]" in failed_log and "error[E0599]" in failed_log,
            "quality-1 visibility diagnostics are missing")
    require("BytesStartExt" in failed_log
            and "checked_attributes" in failed_log
            and "crates/litchi-ole-common/src/xml_attributes.rs" in failed_log
            and "pub(crate) trait BytesStartExt" in failed_log
            and "could not compile `litchi-crypto`" in failed_log,
            "quality-1 failed diagnostic changed")
    return {"archive": "quality-1", "failed_gate": 1,
            "exit_codes": [row["exit_code"] for row in rows],
            "source": {"revision": archived_source["revision"],
                        "files": len(archived_source["files"])},
            "logs": [str(path.relative_to(P)) for path in logs]}


def check_probe_amendment() -> dict[str, Any]:
    """Replay and bind the probe-only setup repairs separately from source."""
    run_reader("probe_amendment_audit.py")
    path = P / "probe-amendment-audit.json"
    value = read(path)
    require(isinstance(value, dict)
            and set(value) == {
                "schema", "passed", "reader", "files", "changes",
                "new_round_trip_test_sha256",
                "counter_logic_and_layout_preserved",
                "fixture_logic_and_tests_preserved",
            }
            and value.get("schema")
            == "litchi.performance.0806.probe-amendment-audit.v1"
            and value.get("passed") is True
            and value.get("counter_logic_and_layout_preserved") is True
            and value.get("fixture_logic_and_tests_preserved") is True,
            "probe amendment audit changed")
    require(value.get("changes") == [
        "test-only allocator registration isolation",
        "feature gate on unused fallback sample helper",
        "seven non-test dead-code allowances on retained support items",
        "unchanged test module relocated to file end",
        "explicit valid-4attr serialization name",
        "six-shape round-trip test with test-only derives",
        "comments and module documentation whitespace",
    ], "probe amendment change inventory changed")
    require(value.get("new_round_trip_test_sha256")
            == "942595b375a60d421c7638d8c8bbd6e9b8283350d1b7c0727471ffb211f81f98",
            "probe amendment round-trip witness changed")
    reader = descriptor(value.get("reader"), "probe amendment reader")
    require(reader is not None and reader.resolve() == (P / "probe_amendment_audit.py").resolve(),
            "probe amendment reader path changed")
    files = value.get("files")
    require(isinstance(files, dict) and set(files) == {
        "main.rs", "allocation_metrics.rs", "counting_allocator.rs"
    }, "probe amendment audit file set changed")
    for name, row in files.items():
        require(isinstance(row, dict) and set(row) == {"initial", "current"},
                f"probe amendment {name} descriptor changed")
        initial = descriptor(row["initial"], f"probe amendment initial {name}")
        current = descriptor(row["current"], f"probe amendment current {name}")
        require(initial is not None and current is not None
                and initial.resolve() == (P / "probe-src-setup-0/src" / name).resolve()
                and current.resolve() == (P / "probe-src/src" / name).resolve(),
                f"probe amendment {name} path changed")
    return {"archive": str(path.relative_to(P)), "passed": True,
            "files": sorted(files), "reader": str(reader.relative_to(P))}


def check_quality_amendment(before: dict[str, Any],
                            candidate_source: dict[str, Any],
                            expected_source: dict[str, Any]) -> dict[str, Any]:
    """Verify the five-helper amendment chain and its applied source witness."""
    root = P / "candidate-quality-amendment"
    manifest_path = root / "manifest.json"
    patch_path = root / "candidate-quality-amendment.patch"
    review_path = root / "source-review.md"
    amendment_inventory = {
        str(path.relative_to(root)) for path in root.rglob("*")
        if path.is_file() and not path.is_symlink()
    }
    require(amendment_inventory == {
        "manifest.json", "candidate-quality-amendment.patch", "source-review.md",
        *(f"before/{name}" for name in AMENDMENT_CHANGES),
        *(f"after/{name}" for name in AMENDMENT_CHANGES),
    }, "quality amendment archive inventory changed")
    manifest = read(manifest_path)
    require(manifest.get("schema") == "litchi.performance.0806.quality-amendment.v1"
            and manifest.get("change") == 806
            and manifest.get("base_commit") == before["revision"],
            "quality amendment manifest identity changed")
    parent = manifest.get("parent_candidate")
    require(isinstance(parent, dict)
            and set(parent) == {"manifest", "patch", "application", "source"},
            "quality amendment parent witness changed")
    expected_parent = {
        "manifest": P / "candidate/manifest.json",
        "patch": P / "candidate/candidate.patch",
        "application": P / "application.json",
        "source": P / "source.json",
    }
    for name, expected_path in expected_parent.items():
        actual = descriptor(parent.get(name), f"quality amendment parent {name}")
        require(actual is not None and actual.resolve() == expected_path.resolve(),
                f"quality amendment parent {name} path changed")
    original_application = read(expected_parent["application"])
    require(original_application.get("source") == candidate_source,
            "original candidate application was altered")
    amendment_files = manifest.get("files")
    require(isinstance(amendment_files, dict)
            and set(amendment_files) == set(AMENDMENT_CHANGES),
            "quality amendment helper archive set changed")
    amended_file_digests = {}
    for archive_name, production in AMENDMENT_CHANGES.items():
        row = amendment_files[archive_name]
        require(isinstance(row, dict)
                and row.get("production_path") == production
                and set(row) == {"production_path", "before", "after"},
                f"quality amendment file record changed: {production}")
        before_path = descriptor(row.get("before"),
                                f"quality amendment before {production}")
        after_path = descriptor(row.get("after"),
                               f"quality amendment after {production}")
        require(before_path is not None and after_path is not None
                and before_path.resolve() == (root / "before" / archive_name).resolve()
                and after_path.resolve() == (root / "after" / archive_name).resolve(),
                f"quality amendment archive path changed: {production}")
        require(c.sha(before_path) == candidate_source["files"][production],
                f"quality amendment before source differs: {production}")
        require(c.sha(after_path) != c.sha(before_path),
                f"quality amendment after source did not change: {production}")
        old = before_path.read_bytes()
        new = after_path.read_bytes()
        constructor = (
            b"    #[allow(clippy::disallowed_methods)]\n"
            b"    #[inline]\n"
            b"    fn new(tag: &'a BytesStart<'a>) -> Self {\n"
            b"        let mut attributes = tag.attributes();\n"
            b"        attributes.with_checks(false);"
        )
        replacement = (
            b"    #[inline]\n"
            b"    fn new(tag: &'a BytesStart<'a>) -> Self {\n"
            b"        let attributes = tag.unchecked_attributes();"
        )
        require(old.count(constructor) == 1 and new == old.replace(constructor, replacement, 1),
                f"quality amendment constructor transformation changed: {production}")
        amended_file_digests[production] = c.sha(after_path)
    shared = manifest.get("shared_files")
    require(isinstance(shared, dict)
            and set(shared) == {"litchi-opc-xml_attributes-tests.rs"},
            "quality amendment shared-file witness changed")
    shared_row = shared["litchi-opc-xml_attributes-tests.rs"]
    require(isinstance(shared_row, dict)
            and set(shared_row) == {"production_path", "original_candidate_after",
                                    "amendment_action"}
            and shared_row.get("production_path") == AMENDMENT_SHARED_TEST
            and shared_row.get("amendment_action") == "byte-identical; omitted from this five-helper amendment",
            "quality amendment shared test record changed")
    shared_archive = descriptor(shared_row.get("original_candidate_after"),
                                "quality amendment shared test")
    require(shared_archive is not None
            and c.sha(shared_archive) == candidate_source["files"][AMENDMENT_SHARED_TEST]
            and shared_archive.resolve() == (P / "candidate/after/litchi-opc-xml_attributes-tests.rs").resolve(),
            "quality amendment shared test changed")
    patch = descriptor(manifest.get("patch"), "quality amendment patch")
    require(patch is not None and patch.resolve() == patch_path.resolve()
            and patch_path.read_text(errors="replace").count("diff --git a/") == 5,
            "quality amendment patch identity changed")
    patch_text = patch_path.read_text(errors="replace")
    require("tests.rs" not in patch_text
            and "litchi-formula" not in patch_text
            and "let attributes = tag.unchecked_attributes();" in patch_text
            and "attributes.with_checks(false);" in patch_text,
            "quality amendment patch scope changed")
    amendment = manifest.get("amendment")
    require(amendment == {
        "reason": "The exact 0805 candidate left CheckedAttributes::unchecked_attributes unused in litchi-sign under warnings-denied probe quality.",
        "operation": "In each of the five helper files, CheckedAttributes::new now obtains its unchecked Attributes iterator through the existing BytesStartExt::unchecked_attributes helper.",
        "removed_allow": "The constructor-local clippy::disallowed_methods allow is removed because the constructor no longer calls quick-xml directly; the helper's existing allow remains in place at its only direct with_checks(false) call.",
        "algorithm_state_and_public_api": "Unchanged. The same unchecked iterator, tag reference, phase initialization, and iterator state are retained; the OPC test source is unchanged.",
        "runtime_requalification_required": True,
        "production_apply_required": True,
    }, "quality amendment rationale changed")
    review = packet_path(manifest.get("review", ""))
    require(review.resolve() == review_path.resolve()
            and review.is_file(), "quality amendment review witness changed")
    application_path = P / "quality-amendment-application.json"
    application = read(application_path)
    require(set(application) == {"schema", "original_application", "manifest",
                                 "patch", "preflight", "source"}
            and application.get("schema")
            == "litchi.performance.0806.quality-amendment-application.v1",
            "quality amendment application schema changed")
    original_descriptor = descriptor(application.get("original_application"),
                                     "quality amendment original application")
    manifest_descriptor = descriptor(application.get("manifest"),
                                     "quality amendment application manifest")
    patch_descriptor = descriptor(application.get("patch"),
                                  "quality amendment application patch")
    preflight_descriptor = descriptor(application.get("preflight"),
                                      "quality amendment application preflight")
    require(original_descriptor is not None
            and original_descriptor.resolve() == (P / "application.json").resolve()
            and manifest_descriptor is not None
            and manifest_descriptor.resolve() == manifest_path.resolve()
            and patch_descriptor is not None
            and patch_descriptor.resolve() == patch_path.resolve()
            and preflight_descriptor is not None
            and preflight_descriptor.resolve() == (P / "amendment-preflight/decision.json").resolve(),
            "quality amendment application artifact chain changed")
    amended_source = application.get("source")
    require(isinstance(amended_source, dict)
            and amended_source.get("revision") == candidate_source["revision"]
            and amended_source.get("files") == expected_source["files"],
            "quality amendment application source differs from intermediate source")
    amended_changes = {
        name for name in set(candidate_source["files"]) | set(amended_source["files"])
        if candidate_source["files"].get(name) != amended_source["files"].get(name)
    }
    require(amended_changes == AMENDMENT_HELPER_FILES,
            "quality amendment changed source outside five helpers")
    require(amended_source["files"][AMENDMENT_SHARED_TEST]
            == candidate_source["files"][AMENDMENT_SHARED_TEST],
            "quality amendment changed shared OPC tests")
    for production, digest in amended_file_digests.items():
        require(amended_source["files"][production] == digest
                and expected_source["files"][production] == digest,
                f"quality amendment source witness differs: {production}")
    return {"manifest": str(manifest_path.relative_to(P)),
            "application": str(application_path.relative_to(P)),
            "changed_files": sorted(amended_changes),
            "source_manifest": amended_source,
            "source": {"revision": amended_source["revision"],
                        "files": len(amended_source["files"])}}


def check_visibility_amendment(before: dict[str, Any],
                               candidate_source: dict[str, Any],
                               quality_source: dict[str, Any],
                               final_source: dict[str, Any]) -> dict[str, Any]:
    """Verify the two-token OLE visibility repair and its application chain."""
    root = P / "candidate-visibility-amendment"
    manifest_path = root / "manifest.json"
    patch_path = root / "candidate-visibility-amendment.patch"
    review_path = root / "source-review.md"
    before_path = root / "before" / VISIBILITY_ARCHIVE
    after_path = root / "after" / VISIBILITY_ARCHIVE
    require(root.is_dir() and not root.is_symlink(),
            "visibility amendment archive is missing")
    inventory = {str(path.relative_to(root)) for path in root.rglob("*")
                 if path.is_file() and not path.is_symlink()}
    require(inventory == {
        "manifest.json", "candidate-visibility-amendment.patch", "source-review.md",
        f"before/{VISIBILITY_ARCHIVE}",
        f"after/{VISIBILITY_ARCHIVE}",
    }, "visibility amendment archive inventory changed")
    manifest = read(manifest_path)
    require(isinstance(manifest, dict)
            and set(manifest) == {
                "schema", "change", "base_commit", "parent_application", "files",
                "patch", "visibility_audit", "scope", "review",
            }
            and manifest.get("schema")
            == "litchi.performance.0806.visibility-amendment.v1"
            and manifest.get("change") == 806
            and manifest.get("base_commit") == before["revision"],
            "visibility amendment manifest identity changed")

    parent_path = P / "quality-amendment-application.json"
    parent_descriptor = descriptor(manifest.get("parent_application"),
                                   "visibility amendment parent application")
    require(parent_descriptor is not None
            and parent_descriptor.resolve() == parent_path.resolve(),
            "visibility amendment parent application path changed")
    parent_application = read(parent_path)
    require(parent_application.get("source") == quality_source,
            "visibility amendment parent source differs from quality amendment")

    files = manifest.get("files")
    require(isinstance(files, dict) and set(files) == {VISIBILITY_ARCHIVE},
            "visibility amendment file set changed")
    row = files[VISIBILITY_ARCHIVE]
    require(isinstance(row, dict)
            and set(row) == {"production_path", "before", "after"}
            and row.get("production_path") == VISIBILITY_PRODUCTION,
            "visibility amendment file record changed")
    archived_before = descriptor(row.get("before"),
                                 "visibility amendment before source")
    archived_after = descriptor(row.get("after"),
                                "visibility amendment after source")
    require(archived_before is not None and archived_after is not None
            and archived_before.resolve() == before_path.resolve()
            and archived_after.resolve() == after_path.resolve(),
            "visibility amendment source archive paths changed")
    require(c.sha(before_path) == quality_source["files"][VISIBILITY_PRODUCTION]
            and c.sha(after_path) == final_source["files"][VISIBILITY_PRODUCTION]
            and c.sha(before_path) != c.sha(after_path),
            "visibility amendment source archive hashes changed")
    old = before_path.read_bytes()
    new = after_path.read_bytes()
    require(old.count(b"pub(crate) trait BytesStartExt") == 1
            and old.count(b"pub(crate) struct CheckedAttributes") == 1,
            "visibility amendment source does not contain the narrowed API")
    expected_new = old.replace(b"pub(crate) trait BytesStartExt",
                               b"pub trait BytesStartExt", 1).replace(
        b"pub(crate) struct CheckedAttributes", b"pub struct CheckedAttributes", 1)
    require(new == expected_new,
            "visibility amendment changed more than the two public tokens")

    patch = descriptor(manifest.get("patch"), "visibility amendment patch")
    require(patch is not None and patch.resolve() == patch_path.resolve(),
            "visibility amendment patch path changed")
    patch_text = patch_path.read_text(errors="replace")
    require(patch_text.count("diff --git a/") == 1
            and patch_text.count("-pub(crate) trait BytesStartExt") == 1
            and patch_text.count("+pub trait BytesStartExt") == 1
            and patch_text.count("-pub(crate) struct CheckedAttributes") == 1
            and patch_text.count("+pub struct CheckedAttributes") == 1
            and "crates/litchi-ole-common/src/xml_attributes.rs" in patch_text
            and "crates/litchi-opc" not in patch_text,
            "visibility amendment patch scope changed")

    audit = manifest.get("visibility_audit")
    helper_names = (
        "litchi-ole-common-xml_attributes.rs",
        "litchi-opc-xml_attributes.rs",
        "litchi-sign-xml_attributes.rs",
        "litchi-xldm-xml_attributes.rs",
        "xml-minifier-xml_attributes.rs",
    )
    helper_paths = {
        "litchi-ole-common-xml_attributes.rs":
            "crates/litchi-ole-common/src/xml_attributes.rs",
        "litchi-opc-xml_attributes.rs": "crates/litchi-opc/src/xml_attributes.rs",
        "litchi-sign-xml_attributes.rs": "crates/litchi-sign/src/xml_attributes.rs",
        "litchi-xldm-xml_attributes.rs": "crates/litchi-xldm/src/xml_attributes.rs",
        "xml-minifier-xml_attributes.rs": "crates/xml-minifier/src/xml_attributes.rs",
    }
    expected_baseline_hashes = {
        name: c.sha(P / "candidate/before" / name) for name in helper_names
    }
    expected_current_hashes = {
        name: quality_source["files"][helper_paths[name]] for name in helper_names
    }
    expected_declarations = {
        "litchi-ole-common-xml_attributes.rs": {
            "baseline": {"BytesStartExt": "pub", "CheckedAttributes": "pub"},
            "current": {"BytesStartExt": "pub(crate)",
                         "CheckedAttributes": "pub(crate)"},
            "status": "two narrowed declarations; amendment restores both",
        },
        "litchi-opc-xml_attributes.rs": {
            "baseline": {"BytesStartExt": "pub", "CheckedAttributes": "pub"},
            "current": {"BytesStartExt": "pub", "CheckedAttributes": "pub"},
            "status": "unchanged",
        },
        "litchi-sign-xml_attributes.rs": {
            "baseline": {"BytesStartExt": "pub(crate)",
                         "CheckedAttributes": "pub(crate)"},
            "current": {"BytesStartExt": "pub(crate)",
                         "CheckedAttributes": "pub(crate)"},
            "status": "unchanged",
        },
        "litchi-xldm-xml_attributes.rs": {
            "baseline": {"BytesStartExt": "pub(crate)",
                         "CheckedAttributes": "pub(crate)"},
            "current": {"BytesStartExt": "pub(crate)",
                         "CheckedAttributes": "pub(crate)"},
            "status": "unchanged",
        },
        "xml-minifier-xml_attributes.rs": {
            "baseline": {"BytesStartExt": "pub(crate)",
                         "CheckedAttributes": "pub(crate)"},
            "current": {"BytesStartExt": "pub(crate)",
                         "CheckedAttributes": "pub(crate)"},
            "status": "unchanged",
        },
    }
    require(audit == {
        "baseline_hashes": expected_baseline_hashes,
        "baseline_source": "docs/performance/results/change-0806/candidate/before/*.rs",
        "current_hashes": expected_current_hashes,
        "current_source": "the five helper files at the parent quality-amendment application boundary",
        "declarations": expected_declarations,
        "result": "The OLE helper's two public declarations are the only visibility narrowings across the five helper files; all other item declarations retain their baseline visibility.",
    }, "visibility amendment audit changed")
    require(manifest.get("scope") == {
        "changed_tokens": 2,
        "changed_files": 1,
        "public_api_restored": [
            "litchi_ole_common::xml_attributes::BytesStartExt",
            "litchi_ole_common::xml_attributes::CheckedAttributes",
        ],
        "algorithm_or_behavior_change": False,
        "runtime_requalification_required": True,
        "production_apply_required": True,
    }, "visibility amendment scope changed")
    review = packet_path(manifest.get("review", ""))
    require(review.resolve() == review_path.resolve() and review.is_file(),
            "visibility amendment review witness changed")

    application_path = P / "visibility-amendment-application.json"
    application = read(application_path)
    require(isinstance(application, dict)
            and set(application) == {"schema", "original_application", "manifest",
                                     "patch", "source"}
            and application.get("schema")
            == "litchi.performance.0806.visibility-amendment-application.v1",
            "visibility amendment application schema changed")
    original_descriptor = descriptor(application.get("original_application"),
                                     "visibility amendment original application")
    manifest_descriptor = descriptor(application.get("manifest"),
                                      "visibility amendment application manifest")
    patch_descriptor = descriptor(application.get("patch"),
                                  "visibility amendment application patch")
    require(original_descriptor is not None
            and original_descriptor.resolve() == parent_path.resolve()
            and manifest_descriptor is not None
            and manifest_descriptor.resolve() == manifest_path.resolve()
            and patch_descriptor is not None
            and patch_descriptor.resolve() == patch_path.resolve(),
            "visibility amendment application artifact chain changed")
    applied_source = application.get("source")
    require(isinstance(applied_source, dict)
            and applied_source == final_source
            and applied_source.get("revision") == before["revision"],
            "visibility amendment application source differs from after build")
    changed = {
        name for name in set(quality_source["files"]) | set(applied_source["files"])
        if quality_source["files"].get(name) != applied_source["files"].get(name)
    }
    require(changed == VISIBILITY_FILES,
            "visibility amendment changed source outside the OLE helper")
    final_changed = {
        name for name in set(before["files"]) | set(applied_source["files"])
        if before["files"].get(name) != applied_source["files"].get(name)
    }
    require(final_changed == SOURCE_ALLOWLIST,
            "visibility-amended source change set differs from frozen allowlist")
    require(candidate_source["files"][VISIBILITY_PRODUCTION]
            != applied_source["files"][VISIBILITY_PRODUCTION],
            "visibility amendment did not preserve candidate narrowing witness")
    return {
        "manifest": str(manifest_path.relative_to(P)),
        "application": str(application_path.relative_to(P)),
        "changed_files": sorted(changed),
        "source_manifest": applied_source,
        "source": {"revision": applied_source["revision"],
                    "files": len(applied_source["files"])},
    }


def check_amendment_preflight(before: dict[str, Any],
                              candidate_source: dict[str, Any],
                              after: dict[str, Any],
                              require_final: bool) -> dict[str, Any]:
    """Check the separate protected native preflight entirely from receipts."""
    root = P / "amendment-preflight"
    def preflight_descriptor(value: Any, label: str) -> Path:
        require(isinstance(value, dict), f"{label} is not an artifact descriptor")
        raw = value.get("path")
        require(isinstance(raw, str) and raw, f"{label}.path is missing")
        path = packet_path(raw) if Path(raw).is_absolute() else (root / raw).resolve()
        require(path.is_file() and not path.is_symlink()
                and path.is_relative_to(root.resolve()),
                f"missing {label}: {raw}")
        actual = c.artifact(path)
        require(actual["bytes"] == value.get("bytes")
                and actual["sha256"] == value.get("sha256"),
                f"{label} identity changed")
        return path
    plan = read(root / "plan.json")
    require(set(plan) == {
                "schema", "purpose", "lineage", "scope", "cpu", "legs", "modes",
                "case_count", "clone_advances", "native", "analysis", "policy",
                "quality", "amendment", "source_allowlist",
            }
            and plan.get("schema") == "litchi.performance.0806.amendment-preflight.v1",
            "amendment preflight plan schema changed")
    require(plan.get("lineage") == {
                "prior_packet": "../change-0805",
                "prior_seal": "../change-0805/seal.json",
                "original_candidate_packet": "../candidate",
                "amendment_handoff": "../candidate-quality-amendment",
                "original_production_before": "../candidate/before",
                "amended_candidate_after": "../candidate-quality-amendment/after",
                "comparison": "source/before versus source/after",
                "historical_timing_pooling": False,
            }, "amendment preflight lineage changed")
    require(plan.get("cpu") == 12
            and plan.get("legs") == ["before", "after"]
            and plan.get("modes") == ["construct", "consume"]
            and plan.get("case_count") == 39,
            "amendment preflight schedule changed")
    require(plan.get("clone_advances") == [0, 1, 2, 3, 4, 5, 32, 33],
            "amendment preflight clone schedule changed")
    native_plan = plan.get("native", {})
    require(native_plan.get("blocks") == 6
            and native_plan.get("samples") == 30
            and native_plan.get("warmup") == 3
            and native_plan.get("iterations") == 4096
            and native_plan.get("orders") == [
                ["before", "after"], ["after", "before"],
                ["before", "after"], ["after", "before"],
                ["after", "before"], ["before", "after"],
            ], "amendment preflight native contract changed")
    analysis_plan = plan.get("analysis", {})
    require(analysis_plan.get("bootstrap_seed") == 806082
            and analysis_plan.get("bootstrap_resamples") == 10_000
            and analysis_plan.get("zero_based_endpoints") == [250, 9749]
            and analysis_plan.get("process_p50") == "nearest rank ceil(n/2)-1",
            "amendment preflight bootstrap contract changed")
    scope = plan.get("scope", {})
    for key in ("public_workflow_speedup", "resource_claim", "production_adoption",
                "profiles", "callgrind", "allocator_measurement"):
        require(scope.get(key) is False,
                f"amendment preflight scope claim changed: {key}")
    require(scope.get("claim") == "protected native micro-input timing only"
            and scope.get("native_process_elapsed_is_diagnostic") is True,
            "amendment preflight claim classification changed")
    quality_plan = plan.get("quality", {})
    require(quality_plan.get("before_test_count") == 70
            and quality_plan.get("after_test_count") == 100
            and quality_plan.get("clippy")
            == "offline locked workspace mirror all targets with -D warnings"
            and quality_plan.get("format") == "cargo fmt --check",
            "amendment preflight quality contract changed")
    require(plan.get("amendment", {}).get("changed_helper_count") == 5
            and plan["amendment"].get("required_constructor_expression")
            == "let attributes = tag.unchecked_attributes();"
            and plan["amendment"].get("forbidden_constructor_expression")
            == "attributes.with_checks(false);"
            and plan["amendment"].get("runtime_algorithm_claim") == "none; only the existing helper call route is repaired"
            and plan["amendment"].get("source_tests_unchanged_from_0805") is True,
            "amendment preflight source contract changed")
    require(plan.get("quality", {}).get("source_paths") == sorted(SOURCE_ALLOWLIST),
            "amendment preflight quality source scope changed")
    require(plan.get("source_allowlist") == sorted(SOURCE_ALLOWLIST),
            "amendment preflight source allowlist changed")
    require(plan.get("policy") == {
        "benefit_mode": "consume",
        "benefit_cases": ["distinct-1", "distinct-2"],
        "benefit_ratio_at_most": 0.97,
        "benefit_ci_high_below": 1.0,
        "protected_consume_cases": [
            "distinct-0", "distinct-1", "distinct-2",
            "duplicate-valid-after-1", "duplicate-valid-after-2",
            "duplicate-long-quoted-after-1", "duplicate-long-quoted-after-2",
            "duplicate-long-quoted-after-33",
            "duplicate-long-unterminated-after-1",
            "duplicate-long-unterminated-after-2",
            "duplicate-long-unterminated-after-33",
            "duplicate-unquoted-after-1", "syntax-flag-after-0",
            "syntax-flag-after-2", "syntax-unique-tail-after-0",
            "syntax-unique-tail-after-2", "syntax-equals-value-after-0",
            "syntax-equals-value-after-2",
        ],
        "protected_ratio_above": 1.05,
        "failure_action": "retain both archives and reject amendment; do not run public workflow adoption captures",
        "success_action": "amendment eligible only for root review; this packet never adopts production",
    }, "amendment preflight policy changed")

    manifest = read(root / "manifest.json")
    require(set(manifest) == {"schema", "packet", "source", "probe", "amendment",
                              "execution", "decision_contract"}
            and manifest.get("schema")
            == "litchi.performance.0806.amendment-preflight-manifest.v1"
            and manifest.get("packet") == "change-0806/amendment-preflight",
            "amendment preflight manifest changed")
    require(manifest.get("probe") == {
                "lineage": "../change-0805/probe-src",
                "case_archive": "../change-0805/cases.json",
                "fixture_archive": "../change-0805/fixtures.json",
                "schema": "litchi.attribute-boundary-probe.v1",
                "tool": "attribute-boundary-probe-0805",
                "clone_advances": [0, 1, 2, 3, 4, 5, 32, 33],
            }
            and manifest.get("amendment") == {
                "source_paths": sorted(AMENDMENT_HELPER_FILES),
                "shared_test_path": AMENDMENT_SHARED_TEST,
                "constructor_rewrite": {
                    "from": "let mut attributes = tag.attributes(); followed by attributes.with_checks(false);",
                    "to": "let attributes = tag.unchecked_attributes();",
                    "algorithm_change": False,
                },
            }, "amendment preflight manifest lineage changed")
    execution = manifest.get("execution", {})
    require(execution.get("cpu") == 12
            and execution.get("native_blocks") == 6
            and execution.get("native_samples") == 30
            and execution.get("native_warmup") == 3
            and execution.get("native_iterations") == 4096
            and execution.get("bootstrap_seed") == 806082
            and execution.get("profiles") is False
            and execution.get("callgrind") is False
            and execution.get("historical_timing_pooling") is False,
            "amendment preflight manifest execution changed")
    require(manifest.get("decision_contract") == {
                "required_fields": [
                    "advance_to_workflow_trials",
                    "production_adoption",
                    "protected_consume_regressions",
                    "dominant_class_benefits",
                ],
                "production_adoption": False,
            }, "amendment preflight decision contract changed")

    handoff = read(root / "handoff.json")
    require(set(handoff) == {"schema", "status", "source", "probe", "drivers",
                             "quality", "native", "exclusions", "production_edits",
                             "cargo_or_native_execution_by_preparer"}
            and handoff.get("schema")
            == "litchi.performance.0806.amendment-preflight-handoff.v1"
            and handoff.get("source") == {
                "before": "source/before", "after": "source/after",
                "files_per_leg": 6,
                "original_before_lineage": "../candidate/before",
                "amended_after_lineage": "../candidate-quality-amendment/after",
            }
            and handoff.get("probe") == {
                "directory": "probe-src", "cases": "cases.json",
                "fixtures": "fixtures.json",
                "source_schema": "litchi.attribute-boundary-probe.v1",
                "tool": "attribute-boundary-probe-0805",
                "clone_advances": [0, 1, 2, 3, 4, 5, 32, 33],
            }
            and handoff.get("drivers") == {
                "plan": "plan.json", "quality": "quality.py",
                "build": "build.py", "capture": "capture.py",
                "reader": "analyze.py", "cleanup": "cleanup.py",
                "custody": "custody.py", "static_audit": "static_audit.py",
            }
            and handoff.get("quality") == {
                "before_tests": 70, "after_tests": 100,
                "clippy": True, "all_targets": True,
            }
            and handoff.get("native") == {
                "build_legs": ["before", "after"],
                "capture_children": 936, "capture_samples": 28080,
                "cpu": 12, "blocks": 6, "samples": 30,
                "warmup": 3, "iterations": 4096, "seed": 806082,
            }
            and handoff.get("exclusions") == {
                "profiles": True, "callgrind": True, "heaptrack": True,
                "historical_timing_pooling": True,
                "reader_in_prebuild_frozen_inputs": True,
            }
            and handoff.get("production_edits") is False
            and handoff.get("cargo_or_native_execution_by_preparer") is False,
            "amendment preflight handoff changed")
    require(handoff.get("status") in {
                "inputs-frozen-execution-pending",
                "inputs-frozen-execution-complete",
                "complete",
            }, "amendment preflight handoff status changed")

    source_names = {
        "litchi-ole-common-xml_attributes.rs": "crates/litchi-ole-common/src/xml_attributes.rs",
        "litchi-opc-xml_attributes.rs": "crates/litchi-opc/src/xml_attributes.rs",
        "litchi-opc-xml_attributes-tests.rs": "crates/litchi-opc/src/xml_attributes/tests.rs",
        "litchi-sign-xml_attributes.rs": "crates/litchi-sign/src/xml_attributes.rs",
        "litchi-xldm-xml_attributes.rs": "crates/litchi-xldm/src/xml_attributes.rs",
        "xml-minifier-xml_attributes.rs": "crates/xml-minifier/src/xml_attributes.rs",
    }
    source_spec = manifest.get("source", {})
    require(source_spec.get("before") == "source/before"
            and source_spec.get("after") == "source/after"
            and source_spec.get("before_lineage") == "../candidate/before"
            and source_spec.get("after_lineage") == "../candidate-quality-amendment/after",
            "amendment preflight source lineage changed")
    preflight_source = source_manifest(root / "source.json",
                                       "amendment preflight workspace source")
    require(preflight_source == candidate_source,
            "amendment preflight workspace source changed")
    for leg, expected_root in (
        ("before", P / "candidate/before"),
        ("after", P / "candidate-quality-amendment/after"),
    ):
        archive = root / "source" / leg
        files = {str(path.relative_to(archive)): path for path in archive.rglob("*")
                 if path.is_file() and not path.is_symlink()}
        require(set(files) == set(source_names),
                f"amendment preflight {leg} source inventory changed")
        for name, production in source_names.items():
            # The quality amendment intentionally archives five helpers.  Its
            # unchanged shared OPC test remains bound to the original
            # candidate-after archive and is copied into both preflight legs.
            lineage_root = (P / "candidate/after"
                            if leg == "after" and name == "litchi-opc-xml_attributes-tests.rs"
                            else expected_root)
            expected = lineage_root / name
            require(c.sha(files[name]) == c.sha(expected),
                    f"amendment preflight {leg} source lineage changed: {name}")
            if leg == "before":
                require(c.sha(files[name]) == before["files"][production],
                        f"amendment preflight before differs from candidate: {production}")
            else:
                require(c.sha(files[name]) == after["files"][production],
                        f"amendment preflight after differs from build: {production}")

    cases = read(root / "cases.json")
    fixtures = read(root / "fixtures.json")
    require(isinstance(cases, list) and len(cases) == 39
            and isinstance(fixtures, list) and len(fixtures) == 39,
            "amendment preflight case/fixture cardinality changed")
    case_ids = [item.get("id") for item in cases]
    require(all(isinstance(item, dict) and isinstance(item.get("id"), str)
                and isinstance(item.get("source"), dict)
                and isinstance(item.get("expected_baseline"), dict)
                for item in cases)
            and len(set(case_ids)) == 39,
            "amendment preflight cases changed")
    fixture_by_id = {item.get("id"): item for item in fixtures}
    require(set(fixture_by_id) == set(case_ids),
            "amendment preflight fixture ids changed")
    for case in cases:
        fixture = fixture_by_id[case["id"]]
        require(set(fixture) == {"id", "category", "attribute_count", "input",
                                "bytes", "sha256"}
                and fixture.get("id") == case["id"]
                and fixture.get("category") == case.get("category")
                and fixture.get("attribute_count") == case.get("attribute_count")
                and isinstance(fixture.get("input"), str)
                and isinstance(fixture.get("bytes"), int)
                and not isinstance(fixture.get("bytes"), bool)
                and fixture.get("bytes") == len(fixture["input"].encode("utf-8"))
                and isinstance(fixture.get("sha256"), str)
                and len(fixture["sha256"]) == 64,
                f"amendment preflight fixture changed: {case['id']}")
        source = case["source"]
        require(isinstance(source, dict)
                and set(source) == {"bytes", "encoding", "value"}
                and source.get("bytes") == fixture["bytes"],
                f"amendment preflight case source changed: {case['id']}")
        if source["encoding"] == "utf8":
            expected_input = source["value"]
        else:
            require(source["encoding"] == "hex"
                    and isinstance(source.get("value"), str),
                    f"amendment preflight source encoding changed: {case['id']}")
            try:
                expected_input = bytes.fromhex(source["value"]).decode("utf-8")
            except (ValueError, UnicodeDecodeError) as error:
                fail(f"amendment preflight source encoding is invalid: {case['id']}: {error}")
        require(fixture["input"] == expected_input
                and fixture["sha256"] == hashlib.sha256(
                    expected_input.encode("utf-8")).hexdigest(),
                f"amendment preflight fixture bytes changed: {case['id']}")

    # Build and mirror quality custody are checked before native reports.  A
    # removed preflight target is accepted only with a separate exact cleanup
    # witness; report custody remains in the packet.
    target_raw = execution.get("target")
    require(isinstance(target_raw, str) and target_raw,
            "amendment preflight target is missing")
    target = Path(target_raw)
    require(target.is_absolute() and target.name == "amendment-preflight",
            "amendment preflight target changed")
    cleanup_path = root / "cleanup.json"
    cleanup = None
    if cleanup_path.is_file():
        cleanup = read(cleanup_path)
        require(cleanup.get("schema") == "litchi.performance.0806.amendment-cleanup.v1"
                and cleanup.get("target") == str(target)
                and cleanup.get("target_removed") is True
                and not target.exists(), "amendment preflight cleanup changed")
        removed = cleanup.get("removed_binaries")
        require(isinstance(removed, list) and len(removed) == 2
                and all(isinstance(item, dict) for item in removed)
                and cleanup.get("removed_failed_binaries") == [],
                "amendment preflight cleanup binary count changed")
        require(isinstance(cleanup.get("removed_target_bytes"), int)
                and cleanup["removed_target_bytes"] >= 0,
                "amendment preflight cleanup byte total changed")
        native_complete = cleanup.get("native_complete")
        require(isinstance(native_complete, dict)
                and native_complete.get("path") == "native/complete.json",
                "amendment preflight cleanup native witness changed")
        preflight_descriptor(native_complete, "amendment preflight cleanup native")
    elif require_final:
        fail("amendment preflight cleanup witness is missing")

    builds = {}
    for leg in ("before", "after"):
        build_root = root / f"build-{leg}"
        build = read(build_root / "build.json")
        expected_archive = {
            name: {"path": f"source/{leg}/{name}",
                   "bytes": (root / "source" / leg / name).stat().st_size,
                   "sha256": c.sha(root / "source" / leg / name)}
            for name in sorted(source_names)
        }
        require(build.get("schema") == "litchi.performance.0806.amendment-build.v1"
                and build.get("leg") == leg
                and build.get("archive") == expected_archive
                and build.get("probe") == {
                    str(path.relative_to(root)): c.sha(path)
                    for path in sorted((root / "probe-src").rglob("*"))
                    if path.is_file() and not path.is_symlink()
                    and path.name != "Cargo.toml"
                }
                and build.get("profiles") is False
                and build.get("callgrind") is False,
                f"amendment preflight {leg} build custody changed")
        frozen = read(build_root / "frozen-inputs.json")
        expected_workspace = candidate_source
        require(frozen.get("schema")
                == "litchi.performance.0806.amendment-build-inputs.v1"
                and frozen.get("leg") == leg
                and frozen.get("plan") == c.sha(root / "plan.json")
                and frozen.get("build_driver") == c.sha(root / "build.py")
                and frozen.get("quality_driver") == c.sha(root / "quality.py")
                and frozen.get("capture_driver") == c.sha(root / "capture.py")
                and frozen.get("custody_driver") == c.sha(root / "custody.py")
                and frozen.get("archives") == {
                    "before": {name: {"path": f"source/before/{name}",
                                      "bytes": (root / "source/before" / name).stat().st_size,
                                      "sha256": c.sha(root / "source/before" / name)}
                                for name in sorted(source_names)},
                    "after": {name: {"path": f"source/after/{name}",
                                     "bytes": (root / "source/after" / name).stat().st_size,
                                     "sha256": c.sha(root / "source/after" / name)}
                               for name in sorted(source_names)},
                }
                and frozen.get("probe") == build.get("probe")
                and frozen.get("workspace_source") == expected_workspace,
                f"amendment preflight {leg} frozen inputs changed")
        binary = build.get("binary")
        require(isinstance(binary, dict)
                and binary.get("path") == str(target / f"{leg}-native")
                and isinstance(binary.get("bytes"), int)
                and isinstance(binary.get("sha256"), str)
                and len(binary["sha256"]) == 64,
                f"amendment preflight {leg} binary descriptor changed")
        binary_path = Path(binary["path"])
        require(binary_path.parent == target and binary_path.name == f"{leg}-native",
                f"amendment preflight {leg} binary escaped target")
        if binary_path.is_file():
            require(c.artifact(binary_path) == binary,
                    f"amendment preflight {leg} binary changed")
        else:
            require(cleanup is not None
                    and any(item == binary for item in cleanup["removed_binaries"]),
                    f"amendment preflight {leg} binary disappeared without cleanup")
        command = build.get("command", {})
        command_args = command.get("command", []) if isinstance(command, dict) else []
        require(isinstance(command, dict) and command.get("exit_code") == 0
                and isinstance(command_args, list)
                and command_args == [
                    "cargo", "build", "--offline", "--locked", "--release",
                    "--manifest-path", str(root / "probe-src/Cargo.toml"),
                    "--bin", "attribute-boundary-probe",
                ],
                f"amendment preflight {leg} build receipt changed")
        build_log = preflight_descriptor(command.get("log"),
                                         f"amendment preflight {leg} build log")
        require(build_log.resolve() == (build_root / "native.log").resolve()
                and isinstance(command.get("started"), (int, float))
                and isinstance(command.get("ended"), (int, float))
                and command["started"] <= command["ended"],
                f"amendment preflight {leg} build timing receipt changed")
        require(build.get("lock") == c.artifact(root / "probe-src/Cargo.lock"),
                f"amendment preflight {leg} lock custody changed")
        require(build.get("environment") == {
                    "CARGO_TARGET_DIR": str(target),
                    "CARGO_BUILD_JOBS": "2",
                    "CARGO_INCREMENTAL": "0",
                    "RUSTFLAGS": None,
                }, f"amendment preflight {leg} build environment changed")
        builds[leg] = build

    if cleanup is not None:
        require({(item.get("path"), item.get("bytes"), item.get("sha256"))
                 for item in cleanup["removed_binaries"]}
                == {(builds[leg]["binary"].get("path"),
                     builds[leg]["binary"].get("bytes"),
                     builds[leg]["binary"].get("sha256"))
                    for leg in ("before", "after")},
                "amendment preflight cleanup binary identities changed")

    quality_root = root / "quality"
    quality_complete = read(quality_root / "complete.json")
    source_archive_hashes = {
        leg: {name: c.sha(root / "source" / leg / name)
              for name in sorted(source_names)}
        for leg in ("before", "after")
    }
    quality_inputs = read(quality_root / "inputs.json")
    source_identity = c.artifact(root / "source.json")
    require(set(quality_inputs) == {"source", "archives", "plan", "driver"}
            and quality_inputs.get("source") == {
                "before": {"path": "source.json",
                            "bytes": source_identity["bytes"],
                            "sha256": source_identity["sha256"]},
                "after": {"path": "source.json",
                           "bytes": source_identity["bytes"],
                           "sha256": source_identity["sha256"]},
            }
            and quality_inputs.get("archives") == source_archive_hashes
            and quality_inputs.get("plan") == c.sha(root / "plan.json")
            and quality_inputs.get("driver") == c.sha(quality_root.parent / "quality.py"),
            "amendment preflight mirror quality inputs changed")
    require(quality_complete.get("schema") == "litchi.performance.0806.amendment-quality.v1"
            and quality_complete.get("test_counts") == {"before": 70, "after": 100}
            and quality_complete.get("expected_test_counts") == {"before": 70, "after": 100}
            and quality_complete.get("scope")
            == "five helper mirror crates plus shared OPC tests; no full production-crate claim"
            and quality_complete.get("source_archives") == source_archive_hashes,
            "amendment preflight mirror quality result changed")
    qrows = read(quality_root / "receipts.json")
    require(isinstance(qrows, list) and len(qrows) == 6
            and quality_complete.get("rows") == qrows,
            "amendment preflight mirror quality receipt count changed")
    previous = None
    expected_legs = ("before", "before", "before", "after", "after", "after")
    expected_commands = ("generate-lockfile", "test", "clippy") * 2
    counts = {}
    for index, row in enumerate(qrows):
        leg = expected_legs[index]
        require(row.get("leg") == leg and row.get("exit_code") == 0
                and row.get("started") <= row.get("ended")
                and (previous is None or previous <= row["started"]),
                f"amendment preflight mirror receipt {index} changed")
        previous = row["ended"]
        command = row.get("command")
        require(isinstance(command, list),
                f"amendment preflight mirror command {index} changed")
        leg_manifest = root / "test-src" / leg / "Cargo.toml"
        expected_command = {
            "generate-lockfile": [
                "cargo", "generate-lockfile", "--offline", "--manifest-path",
                str(leg_manifest),
            ],
            "test": [
                "cargo", "test", "--offline", "--locked", "--manifest-path",
                str(leg_manifest), "--workspace", "--", "--test-threads=2",
            ],
            "clippy": [
                "cargo", "clippy", "--offline", "--locked", "--manifest-path",
                str(leg_manifest), "--workspace", "--all-targets", "--", "-D",
                "warnings",
            ],
        }[expected_commands[index]]
        require(command == expected_command,
                f"amendment preflight mirror command {index} changed")
        log = preflight_descriptor(row.get("log"),
                                   f"amendment preflight mirror log {index}")
        require(log.resolve() == (quality_root / f"{leg}-{index % 3}.log").resolve(),
                f"amendment preflight mirror log {index} escaped")
        if expected_commands[index] == "test":
            matches = re.findall(r"test result: ok\. (\d+) passed; (\d+) failed;",
                                 log.read_text(errors="replace"))
            require(matches and sum(int(item[0]) for item in matches)
                    == {"before": 70, "after": 100}[leg]
                    and sum(int(item[1]) for item in matches) == 0,
                    f"amendment preflight mirror test count changed: {leg}")
            counts[leg] = sum(int(item[0]) for item in matches)
    require(counts == {"before": 70, "after": 100},
            "amendment preflight mirror test evidence missing")

    complete = read(root / "native/complete.json")
    expected_children = 6 * 39 * 2 * 2
    require(complete.get("schema") == "litchi.performance.0806.amendment-native.v1"
            and complete.get("children") == expected_children
            and complete.get("expected_children") == expected_children
            and complete.get("samples") == expected_children * 30
            and complete.get("profiles") is False
            and complete.get("callgrind") is False
            and complete.get("historical_timing_pooling") is False,
            "amendment preflight native completion changed")
    complete_receipts = preflight_descriptor(complete.get("receipts"),
                                             "amendment preflight native receipts")
    complete_source = preflight_descriptor(complete.get("source"),
                                           "amendment preflight native source")
    require(complete_receipts.resolve() == (root / "native/receipts.json").resolve()
            and complete_source.resolve() == (root / "native/source.json").resolve(),
            "amendment preflight native completion artifacts changed")
    capture_source = read(complete_source)
    require(capture_source.get("schema")
            == "litchi.performance.0806.amendment-capture-source.v1"
            and capture_source.get("workspace") == candidate_source
            and capture_source.get("archives") == {
                leg: {name: {
                    "path": f"source/{leg}/{name}",
                    "bytes": (root / "source" / leg / name).stat().st_size,
                    "sha256": c.sha(root / "source" / leg / name),
                } for name in sorted(source_names)}
                for leg in ("before", "after")
            }
            and capture_source.get("probe") == builds["before"].get("probe"),
            "amendment preflight native source witness changed")
    rows = read(root / "native/receipts.json")
    require(isinstance(rows, list) and len(rows) == expected_children,
            "amendment preflight native receipt count changed")
    case_map = {case["id"]: case for case in cases}
    paired: dict[tuple[str, str, int], dict[str, Any]] = {}
    report_paths: set[Path] = set()
    log_paths: set[Path] = set()
    rss_paths: set[Path] = set()
    previous = None
    for index, row in enumerate(rows):
        block = index // (39 * 2 * 2)
        local = index % (39 * 2 * 2)
        case = cases[local // 4]
        mode = plan["modes"][(local % 4) // 2]
        leg = plan["native"]["orders"][block][local % 2]
        label = f"amendment preflight native {index}"
        require(row.get("schema") == "litchi.performance.0806.amendment-native-receipt.v1"
                and row.get("block") == block and row.get("case") == case["id"]
                and row.get("mode") == mode and row.get("leg") == leg
                and row.get("exit_code") == 0
                and row.get("started") <= row.get("ended")
                and (previous is None or previous <= row["started"]),
                f"{label} identity or chronology changed")
        previous = row["ended"]
        require(row.get("binary") == builds[leg]["binary"],
                f"{label} binary binding changed")
        artifact_paths = {}
        for field in ("log", "report", "rss"):
            artifact_path = preflight_descriptor(row.get(field), f"{label}.{field}")
            artifact_paths[field] = artifact_path
            require(artifact_path.resolve().parent == (root / "native").resolve(),
                    f"{label}.{field} escaped native archive")
            if field == "log":
                log_paths.add(artifact_path.resolve())
            elif field == "report":
                report_paths.add(artifact_path.resolve())
            else:
                rss_paths.add(artifact_path.resolve())
        rss = artifact_paths["rss"]
        require(rss.read_text().strip().isdigit()
                and int(rss.read_text().strip()) > 0,
                f"{label} RSS receipt changed")
        command = row.get("command")
        expected_command = [
            "/usr/bin/time", "-f", "%M", "-o", str(rss),
            "taskset", "-c", "12", builds[leg]["binary"]["path"],
            "--leg", leg, "--case", case["id"], "--mode", mode,
            "--samples", "30", "--warmup", "3", "--iterations", "4096",
            "--output", str(artifact_paths["report"]),
        ]
        require(command == expected_command,
                f"{label} command changed")
        report = read(artifact_paths["report"])
        require(report.get("schema") == "litchi.attribute-boundary-probe.v1"
                and report.get("tool") == "attribute-boundary-probe-0805"
                and report.get("binary") == f"{leg}-native"
                and report.get("leg") == leg and report.get("case") == case["id"]
                and report.get("mode") == mode
                and report.get("category") == case.get("category")
                and report.get("attribute_count") == case.get("attribute_count")
                and report.get("iterations") == 4096
                and report.get("warmup") == 3
                and report.get("samples_requested") == 30
                and report.get("source") == case["source"],
                f"{label} report metadata changed")
        timing_scope = {
            "construct": "selected named construction owner call, including its common call dispatch",
            "consume": "selected named consumption owner call, including its common call dispatch",
        }[mode]
        require(report.get("timing_scope") == timing_scope,
                f"{label} timing scope changed")
        oracle = report.get("semantic_oracle")
        require(isinstance(oracle, dict)
                and oracle.get("all_checks_passed") is True
                and oracle.get("baseline") == case["expected_baseline"]
                and oracle.get("candidate") == case["expected_baseline"]
                and oracle.get("baseline_matches_quick_xml") is True
                and oracle.get("candidate_matches_quick_xml") is True
                and oracle.get("quick_xml") == case["expected_baseline"],
                f"{label} semantic oracle changed")
        clones = oracle.get("clone_checks")
        require(isinstance(clones, list) and len(clones) == 8
                and [item.get("advance") for item in clones] == [0, 1, 2, 3, 4, 5, 32, 33]
                and all(item.get("baseline_matches_quick_xml") is True
                            and item.get("candidate_matches_quick_xml") is True
                            and item.get("terminal_behavior_matches") is True
                            for item in clones),
                f"{label} clone oracle changed")
        require(report.get("iterator_sizes")
                == {"baseline_checked_attributes": 120,
                    "candidate_checked_attributes": 128},
                f"{label} iterator layout changed")
        expected_result = report.get("expected_result")
        samples = report.get("samples")
        require(isinstance(expected_result, dict)
                and isinstance(samples, list) and len(samples) == 30,
                f"{label} sample count changed")
        for sample_index, sample in enumerate(samples):
            require(isinstance(sample, dict)
                    and sample.get("index") == sample_index
                    and isinstance(sample.get("elapsed_ns"), int)
                    and not isinstance(sample.get("elapsed_ns"), bool)
                    and sample["elapsed_ns"] > 0
                    and all(sample.get(field) == expected_result.get(field)
                            for field in ("checksum", "accepted", "error_marker")),
                    f"{label} sample {sample_index} changed")
        paired.setdefault((case["id"], mode, block), {})[leg] = report

    require(len(report_paths) == expected_children
            and len(log_paths) == expected_children
            and len(rss_paths) == expected_children,
            "amendment preflight native artifact paths are duplicated")

    # Recompute the protected analysis from the retained samples and bind the
    # decision.  This keeps the supplemental lane diagnostic and prevents a
    # forged decision from promoting an unverified report set.
    analysis = read(root / "analysis.json")
    require(analysis.get("schema") == "litchi.performance.0806.amendment-analysis.v1"
            and analysis.get("policy") == plan["policy"]
            and analysis.get("claims") == plan["scope"],
            "amendment preflight analysis contract changed")
    analysis_rows = analysis.get("rows")
    require(isinstance(analysis_rows, list) and len(analysis_rows) == 39 * 2,
            "amendment preflight analysis row count changed")
    def p50(values: list[int]) -> float:
        ordered = sorted(values)
        return ordered[(len(ordered) + 1) // 2 - 1]
    def bootstrap(values: list[float]) -> tuple[float, float, float]:
        rng = random.Random(806082)
        medians = [statistics.median(values[rng.randrange(len(values))]
                                   for _ in values)
                   for _ in range(10_000)]
        medians.sort()
        return statistics.median(values), medians[250], medians[9749]
    computed = {}
    for case in cases:
        for mode in plan["modes"]:
            before_p50 = []
            after_p50 = []
            ratios = []
            for block in range(6):
                pair = paired[(case["id"], mode, block)]
                require(set(pair) == {"before", "after"},
                        f"analysis pair missing: {case['id']}/{mode}/{block}")
                before_value = p50([sample["elapsed_ns"] for sample in pair["before"]["samples"]])
                after_value = p50([sample["elapsed_ns"] for sample in pair["after"]["samples"]])
                require(before_value > 0, f"zero protected before p50: {case['id']}/{mode}/{block}")
                before_p50.append(before_value)
                after_p50.append(after_value)
                ratios.append(after_value / before_value)
            ratio, low, high = bootstrap(ratios)
            computed[(case["id"], mode)] = {
                "case": case["id"], "mode": mode,
                "process_p50_before": before_p50,
                "process_p50_after": after_p50,
                "paired_ratios": ratios,
                "ratio_median": ratio,
                "bootstrap_ci_low": low,
                "bootstrap_ci_high": high,
                "change_percent_median": (ratio - 1.0) * 100.0,
                "diagnostic_regression": ratio > 1.05 and low > 1.0,
            }
    require([(row.get("case"), row.get("mode")) for row in analysis_rows]
            == [(case["id"], mode) for case in cases for mode in plan["modes"]],
            "amendment preflight analysis order changed")
    for row in analysis_rows:
        expected = computed[(row["case"], row["mode"])]
        require(row == expected,
                f"amendment preflight analysis row changed: {row['case']}/{row['mode']}")
    decision = read(root / "decision.json")
    require(set(decision) == {
                "schema", "legacy_schema", "packet", "seed",
                "advance_to_workflow_trials", "production_adoption",
                "protected_consume_regressions", "all_consume_regressions",
                "dominant_class_benefits", "benefit_policy_passed",
                "independent_audit", "counts", "analysis", "custody",
            }
            and decision.get("schema") == "litchi.performance.0806.amendment-decision.v1"
            and decision.get("legacy_schema")
            == "litchi.performance.0805.preflight-decision.v1"
            and decision.get("packet") == "change-0806/amendment-preflight"
            and decision.get("seed") == 806082
            and decision.get("advance_to_workflow_trials") is True
            and decision.get("production_adoption") is False
            and decision.get("protected_consume_regressions") == []
            and decision.get("dominant_class_benefits") == {"distinct-1": True, "distinct-2": True}
            and decision.get("benefit_policy_passed") is True
            and decision.get("all_consume_regressions") == [
                {
                    "case": row["case"], "mode": row["mode"],
                    "ratio_median": row["ratio_median"],
                    "bootstrap_ci_low": row["bootstrap_ci_low"],
                    "bootstrap_ci_high": row["bootstrap_ci_high"],
                    "change_percent_median": row["change_percent_median"],
                }
                for row in analysis_rows
                if row["mode"] == "consume" and row["diagnostic_regression"] is True
            ],
            "amendment preflight decision changed")
    require(decision.get("counts") == {
        "case_count": 39, "native_reports": 936,
        "native_children": 936, "native_samples": 28_080,
    }, "amendment preflight counts changed")
    analysis_descriptor = preflight_descriptor(decision.get("analysis"),
                                               "amendment preflight analysis")
    require(analysis_descriptor.resolve() == (root / "analysis.json").resolve()
            and read(analysis_descriptor) == analysis,
            "amendment preflight decision analysis differs")
    audit_descriptor = preflight_descriptor(decision.get("independent_audit"),
                                             "amendment preflight root native audit")
    require(audit_descriptor.resolve() == (root / "root-native-audit.json").resolve(),
            "amendment preflight audit witness changed")
    run_reader("amendment-preflight/root_native_audit.py")
    audit = read(audit_descriptor)
    require(audit.get("schema") == "litchi.performance.0806.amendment-root-native-audit.v1"
            and audit.get("passed") is True
            and audit.get("matches_primary_analysis") is True
            and audit.get("native_reports") == 936
            and audit.get("native_samples") == 28_080
            and audit.get("advance_to_workflow_trials") is True
            and audit.get("production_adoption") is False
            and audit.get("protected_consume_regressions") == []
            and audit.get("dominant_class_benefits") == {
                "distinct-1": True, "distinct-2": True,
            }, "amendment preflight root audit changed")
    custody = decision.get("custody", {})
    require(set(custody) == {"archives", "builds", "native", "probe", "source",
                             "quality", "profiles", "callgrind",
                             "historical_timing_pooling"}
            and custody.get("archives") == {
                leg: {name: {
                    "path": f"source/{leg}/{name}",
                    "bytes": (root / "source" / leg / name).stat().st_size,
                    "sha256": c.sha(root / "source" / leg / name),
                } for name in sorted(source_names)}
                for leg in ("before", "after")
            }
            and custody.get("profiles") is False
            and custody.get("callgrind") is False
            and custody.get("historical_timing_pooling") is False,
            "amendment preflight decision custody claims changed")
    for leg in ("before", "after"):
        build_descriptor = preflight_descriptor(custody["builds"].get(leg),
                                                f"amendment preflight decision {leg} build")
        require(build_descriptor.resolve() == (root / f"build-{leg}/build.json").resolve(),
                f"amendment preflight decision {leg} build path changed")
    require(preflight_descriptor(custody["native"],
                                 "amendment preflight decision native").resolve()
            == (root / "native/complete.json").resolve()
            and preflight_descriptor(custody["source"],
                                     "amendment preflight decision source").resolve()
            == (root / "source.json").resolve()
            and preflight_descriptor(custody["quality"],
                                     "amendment preflight decision quality").resolve()
            == (root / "quality/complete.json").resolve(),
            "amendment preflight decision custody paths changed")
    require(custody.get("probe") == builds["before"].get("probe"),
            "amendment preflight decision probe custody changed")
    return {"cases": 39, "native_reports": 468, "native_children": 936,
            "native_samples": 28_080, "test_counts": {"before": 70, "after": 100},
            "advance_to_workflow_trials": True, "production_adoption": False,
            "cleanup": cleanup is not None}


def check_probe_quality(before: dict[str, Any], after: dict[str, Any]) -> dict[str, Any]:
    summaries = {}
    for leg, expected_source in (("before", before), ("after", after)):
        root = P / f"probe-quality-{leg}"
        complete = read(root / "complete.json")
        inputs = descriptor(complete.get("inputs"), f"probe-quality-{leg}.inputs")
        require(inputs is not None, "probe quality inputs missing")
        input_value = read(inputs)
        source = input_value.get("source")
        require(isinstance(source, dict) and source == expected_source,
                f"probe quality {leg} embedded source census changed")
        require(isinstance(input_value.get("probe"), dict),
                f"probe quality {leg} probe inventory missing")
        driver = descriptor(input_value.get("driver"), f"probe-quality-{leg}.driver")
        require(driver is not None, "probe quality driver missing")
        receipts_path = descriptor(complete.get("receipts"),
                                   f"probe-quality-{leg}.receipts")
        require(receipts_path is not None, "probe quality receipts missing")
        rows = read(receipts_path)
        require(isinstance(rows, list) and len(rows) == 3,
                f"probe quality {leg} gate count changed")
        previous = None
        test_summary = None
        for index, row in enumerate(rows):
            require(row.get("exit_code") == 0 and row.get("started") <= row.get("ended"),
                    f"probe quality {leg} gate {index} failed")
            if previous is not None:
                require(previous <= row["started"], "probe quality gates not serial")
            previous = row["ended"]
            log = descriptor(row.get("log"), f"probe-quality-{leg}.{index}.log")
            require(log is not None, "probe quality log missing")
            if index == 1:
                matches = re.findall(
                    r"^test result: (?:ok|FAILED)\.\s+(\d+) passed; (\d+) failed; "
                    r"(\d+) ignored; (\d+) measured; (\d+) filtered out;",
                    log.read_text(errors="replace"), re.MULTILINE)
                require(matches, f"probe quality {leg} test log has no result")
                test_summary = {"suites": len(matches),
                                "passed": sum(int(item[0]) for item in matches),
                                "failed": sum(int(item[1]) for item in matches),
                                "ignored": sum(int(item[2]) for item in matches)}
                require(test_summary["failed"] == 0,
                        f"probe quality {leg} contains a failed suite")
        summaries[leg] = test_summary
    return {"gates": 6, "before": summaries["before"], "after": summaries["after"]}


def check_main_receipts(builds: dict[str, Any], cleanup: Any,
                        cleanup_ok: bool) -> dict[str, Any]:
    plan = read(P / "plan.json")
    require(plan.get("schema") == "litchi.performance.0806.v1",
            "main plan schema changed")
    for lane, blocks, samples, warmup, expected_count, expected_source in (
        ("qualification", 1, 1, 0, QUALIFICATION_CHILDREN, "before"),
        ("native", 6, 30, 3, NATIVE_CHILDREN, "after"),
        ("allocation", 2, 3, 0, ALLOCATION_CHILDREN, "after"),
    ):
        root = P / lane
        complete = read(root / "complete.json")
        source_path, source = source_descriptor(complete.get("source"), f"{lane}.source")
        build_source = source_manifest(Path(builds[expected_source]["source"]["path"]),
                                       f"build-{expected_source}.source")
        require(source == build_source, f"{lane} source differs from build")
        receipts_path = descriptor(complete.get("receipts"), f"{lane}.receipts")
        require(receipts_path is not None, f"{lane} receipts missing")
        rows = read(receipts_path)
        require(isinstance(rows, list) and len(rows) == expected_count
                and complete.get("children") == expected_count,
                f"{lane} receipt cardinality changed")
        expected = []
        for block in range(blocks):
            order = ("before",) if lane == "qualification" else ORDERS[block]
            for shape, mode in CASES:
                for leg in order:
                    expected.append((block, shape, mode, leg))
        require(len(expected) == expected_count, f"{lane} schedule changed")
        previous = None
        for index, (row, identity) in enumerate(zip(rows, expected)):
            block, shape, mode, leg = identity
            require((row.get("lane"), row.get("block"), row.get("shape"),
                     row.get("mode"), row.get("leg"))
                    == (lane, block, shape, mode, leg),
                    f"{lane} row {index} identity changed")
            require(row.get("exit_code") == 0 and row.get("started") <= row.get("ended"),
                    f"{lane} row {index} failed")
            if previous is not None:
                require(previous <= row["started"], f"{lane} receipts not serial")
            previous = row["ended"]
            kind = "allocation" if lane in ("allocation", "qualification") else "native"
            binary = builds[leg]["binaries"][kind]
            require(row.get("binary") == binary, f"{lane} row {index} binary changed")
            descriptor(binary, f"{lane} row {index}.binary", allow_missing=cleanup_ok,
                       external=True)
            log = descriptor(row.get("log"), f"{lane} row {index}.log")
            report = descriptor(row.get("report"), f"{lane} row {index}.report")
            rss = descriptor(row.get("rss"), f"{lane} row {index}.rss")
            require(log is not None and report is not None and rss is not None,
                    f"{lane} row {index} artifact missing")
            require(rss.read_text().strip().isdigit(), f"{lane} row {index} RSS is invalid")
            command = row.get("command")
            require(isinstance(command, list) and command[0:4]
                    == ["/usr/bin/time", "-f", "%M", "-o"],
                    f"{lane} row {index} command wrapper changed")
            require("taskset" in command and "-c" in command
                    and command[command.index("-c") + 1] == "12",
                    f"{lane} row {index} CPU changed")
            require("--mode" in command and command[command.index("--mode") + 1] == mode
                    and "--shape" in command and command[command.index("--shape") + 1] == shape,
                    f"{lane} row {index} fixture command changed")
            require("--samples" in command
                    and command[command.index("--samples") + 1] == str(samples)
                    and "--warmup" in command
                    and command[command.index("--warmup") + 1] == str(warmup),
                    f"{lane} row {index} sample command changed")
            require(Path(command[command.index("--output") + 1]).resolve() == report.resolve(),
                    f"{lane} row {index} report command path changed")
            raw = read(report)
            require(isinstance(raw, dict) and isinstance(raw.get("samples"), list)
                    and len(raw["samples"]) == samples,
                    f"{lane} row {index} report sample count changed")
        require(source_path.is_file(), f"{lane} source disappeared")
    return {"native_reports": NATIVE_CHILDREN, "native_samples": NATIVE_SAMPLES,
            "allocation_reports": ALLOCATION_CHILDREN,
            "allocation_samples": ALLOCATION_SAMPLES,
            "qualification_reports": QUALIFICATION_CHILDREN,
            "qualification_samples": QUALIFICATION_SAMPLES}


def check_profile() -> dict[str, Any]:
    run_reader("profile_analysis.py")
    value = read(P / "profile-analysis.json")
    require(isinstance(value, dict)
            and value.get("reports") == PROFILE_REPORTS
            and value.get("samples") == PROFILE_SAMPLES
            and value.get("decodes") == 8
            and isinstance(value.get("rows"), list)
            and len(value["rows"]) == PROFILE_REPORTS,
            "profile aggregate changed")
    require(value.get("timing_claim") is False and value.get("rss_claim") is False,
            "profile timing/RSS claim changed")
    for row in value["rows"]:
        require(row.get("qualified") is all(row.get("checks", {}).values()),
                "profile qualification changed")
    return {"reports": PROFILE_REPORTS, "samples": PROFILE_SAMPLES,
            "decodes": 8, "qualified": value.get("qualified")}


def check_root_audit() -> dict[str, Any]:
    run_reader("root_audit.py")
    value = read(P / "root-audit.json")
    require(value.get("schema") == "litchi.performance.0806.root-audit.v1"
            and value.get("reports") == MAIN_REPORTS
            and value.get("samples") == MAIN_SAMPLES
            and len(value.get("native", [])) == len(CASES)
            and len(value.get("allocation", [])) == 2 * len(CASES)
            and {(row.get("shape"), row.get("mode"), row.get("block"))
                 for row in value.get("allocation", [])}
            == {(shape, mode, block) for shape, mode in CASES
                for block in (0, 1)}
            and len(value.get("qualification", [])) == len(CASES),
            "independent root audit aggregate changed")
    require(isinstance(value.get("latency_violations"), list)
            and isinstance(value.get("resource_violations"), list)
            and isinstance(value.get("benefits"), list)
            and value.get("resource_guard") is (not value["resource_violations"]),
            "independent root audit guard fields changed")
    require(value.get("production_adoption") is False,
            "independent root audit made an adoption decision")
    return {"reports": MAIN_REPORTS, "samples": MAIN_SAMPLES,
            "latency_violations": len(value.get("latency_violations", [])),
            "resource_violations": len(value.get("resource_violations", [])),
            "benefits": len(value.get("benefits", [])),
            "passed": not value.get("latency_violations", [])
            and not value.get("resource_violations", [])}


def check_quality_summary() -> dict[str, Any]:
    run_reader("quality_summary.py")
    value = read(P / "quality-summary.json")
    require(value.get("schema") == "litchi.performance.0806.quality-summary.v1"
            and value.get("gates") == 6
            and value.get("failed") == 0
            and value.get("suites", 0) > 0,
            "quality summary aggregate changed")
    return {"gates": value["gates"], "suites": value["suites"],
            "passed": value["passed"], "failed": value["failed"]}


def check_cross() -> dict[str, Any]:
    run_reader("cross_analysis.py")
    run_reader("cross_root_audit.py")
    value = read(P / "cross-analysis.json")
    require(value.get("schema") == "litchi.performance.0806.cross-format-analysis.v1",
            "cross analysis schema changed")
    require(value.get("counts") == {
                "native_reports": 12,
                "qualification_reports": 1,
                "native_samples": 2_880,
                "qualification_samples": 8,
                "analysis_rows": 8,
            }, "cross cardinality changed")
    require(len(value.get("rows", [])) == 8
            and isinstance(value.get("decision", {}).get("rejected"), bool),
            "cross disposition is not boolean")
    root = read(P / "cross-root-audit.json")
    require(root.get("raw_process_rows") == 96 and len(root.get("rows", [])) == 8
            and root.get("rejected") == value["decision"]["rejected"],
            "cross independent audit disagrees")
    cleanup = read(P / "cross-cleanup.json")
    require(set(cleanup) == {"schema", "target", "binaries_removed", "removed_binaries"}
            and cleanup.get("schema") == "litchi.performance.0806.cross-cleanup.v1"
            and cleanup.get("target") == str(TARGET)
            and cleanup.get("binaries_removed") is True,
            "cross cleanup changed")
    removed = cleanup.get("removed_binaries")
    require(isinstance(removed, list) and len(removed) == 2,
            "cross cleanup binary count changed")
    expected_paths = {str(TARGET / "cross-before"), str(TARGET / "cross-after")}
    for item in removed:
        require(isinstance(item, dict)
                and isinstance(item.get("path"), str)
                and item["path"] in expected_paths
                and not Path(item["path"]).exists()
                and isinstance(item.get("bytes"), int)
                and len(item.get("sha256", "")) == 64,
                "cross cleanup identity malformed or escaped target")
    require({item["path"] for item in removed} == expected_paths,
            "cross cleanup binary names changed")
    return {"reports": CROSS_REPORTS, "samples": CROSS_SAMPLES,
            "native_reports": 12, "qualification_reports": 1,
            "native_samples": 2_880, "qualification_samples": 8,
            "rejected": value["decision"]["rejected"]}


def check_cross_qualification() -> dict[str, Any]:
    """Check the before-only cross-format corpus identity witness.

    The writer is intentionally not rerun here: it is a root capture witness,
    and the retained JSON is checked against its exact artifact identities.
    """
    value = read(P / "cross-qualification-audit.json")
    require(value.get("schema") ==
            "litchi.performance.0806.cross-qualification-audit.v1"
            and value.get("passed") is True
            and value.get("reports") == 1
            and value.get("samples") == 8
            and value.get("historical_timings_imported") is False,
            "cross qualification audit changed")
    for name in ("build_receipt", "report", "historical_fixture_reference",
                 "historical_seal", "source", "reader"):
        descriptor(value.get(name), f"cross qualification {name}",
                   external=name in {"historical_fixture_reference", "historical_seal"})
    matches = value.get("fixture_identity_matches")
    require(isinstance(matches, list) and len(matches) == 8,
            "cross qualification corpus cardinality changed")
    return {"reports": 1, "samples": 8, "fixture_rows": len(matches)}


def check_analysis() -> dict[str, Any]:
    run_reader("analyze.py")
    value = read(P / "analysis.json")
    require(isinstance(value, dict)
            and isinstance(value.get("schema"), str)
            and ".0806." in value["schema"]
            and value.get("plan_schema") == "litchi.performance.0806.v1",
            "main analysis schema changed")
    counts = value.get("counts")
    require(counts == {"reports": MAIN_REPORTS, "samples": MAIN_SAMPLES,
                       "native_reports": NATIVE_CHILDREN,
                       "allocation_reports": ALLOCATION_CHILDREN,
                       "qualification_reports": QUALIFICATION_CHILDREN},
            f"main aggregate counts changed: {counts}")
    require(value.get("native", {}).get("children") == NATIVE_CHILDREN
            and value.get("allocation", {}).get("children") == ALLOCATION_CHILDREN
            and value.get("qualification", {}).get("children") == QUALIFICATION_CHILDREN,
            "main lane cardinality changed")
    verification = value.get("verification", {})
    for key in ("all_raw_report_samples_checked", "semantic_verification_checked",
                "source_binary_probe_lock_receipts_checked", "frozen_inputs_checked",
                "architecture_inputs_checked", "quality_commands_checked_exactly",
                "allocation_memory_guards_use_block_medians",
                "allocation_calls_and_bytes_guards_checked",
                "historical_qualification_checked_without_timings",
                "new_four_attribute_oracle_checked_without_timings",
                "after_build_inputs_checked", "probe_tests_checked",
                "quality_amendment_application_checked",
                "visibility_amendment_application_checked",
                "amendment_preflight_checked"):
        require(verification.get(key) is True, f"analysis verification missing: {key}")
    source = value.get("source", {})
    require(set(source.get("changed_files", ())) == SOURCE_ALLOWLIST,
            "analysis source change set changed")
    guards = value.get("decision_guards", {})
    require(isinstance(guards.get("latency_guard_passed"), bool)
            and isinstance(guards.get("resource_guard_passed"), bool)
            and isinstance(guards.get("benefit_satisfied"), bool),
            "analysis decision guards malformed")
    return value


def check_disposition(before: dict[str, Any], after: dict[str, Any],
                      analysis: dict[str, Any], quality: dict[str, Any],
                      quality_summary: dict[str, Any], probe_quality: dict[str, Any],
                      profile: dict[str, Any], root_audit: dict[str, Any],
                      cross: dict[str, Any], quality_failure: dict[str, Any],
                      quality_failure_one: dict[str, Any],
                      probe_amendment: dict[str, Any],
                      quality_amendment: dict[str, Any],
                      visibility_amendment: dict[str, Any],
                      amendment_preflight: dict[str, Any],
                      require_final: bool) -> dict[str, Any]:
    path = P / "disposition.json"
    if not path.is_file():
        require(not require_final, "final disposition is missing")
        require(current_source() == after["files"],
                "pre-disposition source is not the candidate")
        return {"status": "pending", "final": False}
    value = read(path)
    status = value.get("status")
    require(status in {"retained", "rejected"}, "disposition status changed")
    require(value.get("production_change_retained") is (status == "retained"),
            "disposition adoption flag changed")
    gates = {
        "main_decision_guards": analysis["decision_guards"].get("adoption_eligible") is True,
        "quality": quality.get("gates") == 6,
        "quality_summary": quality_summary.get("failed") == 0,
        "probe_quality": (probe_quality.get("before", {}).get("failed", 0) == 0
                          and probe_quality.get("after", {}).get("failed", 0) == 0),
        "profile": profile.get("qualified") is True,
        "main_root_audit": root_audit.get("passed") is True,
        "main_root_benefit": root_audit.get("benefits", 0) > 0,
        "cross": cross.get("rejected") is False,
        "quality_failure_retained": quality_failure.get("failed_gate") == 1,
        "quality_failure_one_retained": quality_failure_one.get("failed_gate") == 1,
        "probe_amendment": probe_amendment.get("passed") is True,
        "quality_amendment": bool(quality_amendment.get("changed_files")),
        "visibility_amendment": bool(visibility_amendment.get("changed_files")),
        "amendment_preflight": (
            amendment_preflight.get("advance_to_workflow_trials") is True
            and amendment_preflight.get("production_adoption") is False
        ),
    }
    failed_gates = sorted(name for name, passed in gates.items() if not passed)
    expected = after if status == "retained" else before
    require(current_source() == expected["files"],
            "live source does not match disposition")
    if status == "rejected":
        require(failed_gates, "rejected disposition has no failed adoption gate")
        require(isinstance(value.get("reason"), str) and value["reason"].strip(),
                "rejected disposition has no actual reason")
        restored = descriptor(value.get("restored_source"), "restored source")
        require(restored is not None and source_manifest(restored, "restored source")["files"]
                == before["files"], "restored source witness changed")
    if status == "retained":
        require(not failed_gates, f"retained disposition has failed gates: {failed_gates}")
    return {"status": status, "final": True, "gates": gates,
            "failed_gates": failed_gates}


def check_seal(analysis: dict[str, Any], disposition: dict[str, Any]) -> dict[str, Any]:
    seal = read(P / "seal.json")
    require(seal.get("schema") == "litchi.performance.0806.final-seal.v1",
            "final seal schema changed")
    actual = {str(path.relative_to(P)): c.sha(path)
              for path in P.rglob("*")
              if path.is_file() and path.name != "seal.json" and "__pycache__" not in path.parts}
    require(seal.get("files") == actual, "final packet seal inventory is stale")
    docs = seal.get("documents")
    require(isinstance(docs, dict) and len(docs) == 6,
            "final document seal cardinality changed")
    for name, digest in docs.items():
        path = ROOT / name
        require(path.is_file() and c.sha(path) == digest,
                f"sealed document changed: {name}")
    production = seal.get("production")
    expected = analysis["source"]["after"]["files"] if disposition["status"] == "retained" else {}
    baseline = analysis["source"]["before"]["files"]
    expected = {name: digest for name, digest in expected.items()
                if digest != baseline.get(name)}
    require(production == expected, "final production seal disposition changed")
    return {"files": len(actual), "documents": len(docs), "production": len(production)}


def validate(require_final: bool, allow_precleanup: bool,
             check_worktree: bool) -> dict[str, Any]:
    workspace = check_workspace() if (check_worktree or require_final) else None
    origin = read(P / "origin.json")
    before = source_manifest(P / "source.json", "production source")
    require(before["revision"] == origin["base"], "production source revision changed")
    build_before = read(P / "build-before/build.json")
    build_after = read(P / "build-after/build.json")
    _, before_build = source_descriptor(build_before.get("source"), "build-before.source")
    _, after = source_descriptor(build_after.get("source"), "build-after.source")
    require(before_build == before, "before build differs from frozen production")
    setup = check_setup_archives(before, require_final)
    original_application = read(P / "application.json")
    original_source = original_application.get("source")
    require(isinstance(original_source, dict)
            and isinstance(original_source.get("revision"), str)
            and isinstance(original_source.get("files"), dict),
            "original candidate source witness is malformed")
    candidate_source = {"revision": original_source["revision"],
                        "files": dict(original_source["files"])}
    candidate = check_candidate(before, candidate_source)
    cleanup, cleanup_ok = cleanup_records()
    if not cleanup_ok:
        require(allow_precleanup, "owned target is live; pass --allow-precleanup")
        require(TARGET.exists(), "owned target is missing without cleanup witness")
    builds = check_builds(before, after, cleanup, cleanup_ok)
    candidate_source = candidate.pop("candidate_source")
    quality_failure = check_quality_failure(candidate_source)
    probe_amendment = check_probe_amendment()
    quality_application = read(P / "quality-amendment-application.json")
    quality_source = quality_application.get("source")
    require(isinstance(quality_source, dict)
            and isinstance(quality_source.get("revision"), str)
            and isinstance(quality_source.get("files"), dict),
            "quality amendment source witness is malformed")
    quality_amendment = check_quality_amendment(
        before, candidate_source, quality_source
    )
    quality_source = quality_amendment.pop("source_manifest")
    quality_failure_one = check_quality_failure_one(quality_source)
    amendment_preflight = check_amendment_preflight(
        before, candidate_source, quality_source, require_final
    )
    visibility_amendment = check_visibility_amendment(
        before, candidate_source, quality_source, after
    )
    visibility_amendment.pop("source_manifest")
    quality = check_quality(after)
    quality_summary = check_quality_summary()
    probe_quality = check_probe_quality(before, after)
    main = check_main_receipts(builds, cleanup, cleanup_ok)
    profile = check_profile()
    root_audit = check_root_audit()
    cross_qualification = check_cross_qualification()
    cross = check_cross()
    analysis = check_analysis()
    disposition = check_disposition(before, after, analysis, quality,
                                    quality_summary, probe_quality, profile,
                                    root_audit, cross, quality_failure,
                                    quality_failure_one,
                                    probe_amendment, quality_amendment,
                                    visibility_amendment,
                                    amendment_preflight, require_final)
    if require_final:
        require(cleanup_ok, "final main cleanup witness is missing")
        seal = check_seal(analysis, disposition)
    else:
        seal = None
    if require_final:
        require((P / "failure-audit.json").is_file(),
                "final failure audit is missing")
    if (P / "failure_audit.py").is_file() and (P / "failure-audit.json").is_file():
        run_reader("failure_audit.py")
    require(not list(P.rglob("__pycache__")), "packet contains __pycache__")
    return {
        "main_reports": MAIN_REPORTS,
        "main_samples": MAIN_SAMPLES,
        "profile_reports": PROFILE_REPORTS,
        "profile_samples": PROFILE_SAMPLES,
        "cross_reports": CROSS_REPORTS,
        "cross_samples": CROSS_SAMPLES,
        "total_reports": TOTAL_REPORTS,
        "total_samples": TOTAL_SAMPLES,
        "quality": quality,
        "quality_failure": quality_failure,
        "quality_failure_one": quality_failure_one,
        "probe_amendment": probe_amendment,
        "quality_amendment": quality_amendment,
        "visibility_amendment": visibility_amendment,
        "amendment_preflight": amendment_preflight,
        "quality_summary": quality_summary,
        "probe_quality": probe_quality,
        "candidate": candidate,
        "setup": setup,
        "main": main,
        "profile": profile,
        "root_audit": root_audit,
        "cross_qualification": cross_qualification,
        "cross": cross,
        "disposition": disposition,
        "seal_checked": seal is not None,
        "workspace_checked": workspace is not None,
    }


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--require-final-seal", action="store_true")
    parser.add_argument("--allow-precleanup", action="store_true")
    parser.add_argument("--check-workspace", action="store_true")
    args = parser.parse_args()
    try:
        print(json.dumps(validate(args.require_final_seal, args.allow_precleanup,
                                  args.check_workspace), indent=2, sort_keys=True))
    except (AssertionError, subprocess.CalledProcessError) as error:
        print(f"validation failed: {error}", file=sys.stderr)
        raise SystemExit(1)
