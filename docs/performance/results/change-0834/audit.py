#!/usr/bin/env python3
"""Independent custody and chronology audit for the 0834 repair packet.

This reader owns packet custody, command receipt identity, serial chronology,
and the retained diagnostic's raw-read arithmetic.  It does not import the
admission reader or the diagnostic analyzer, and it never launches a workload,
build, Git command, or external parser.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import random
import statistics
import sys
from pathlib import Path
from typing import Any, Iterable

import driver


PACKET = Path(driver.P).resolve()
ROOT = Path(driver.ROOT).resolve()
TARGET = Path(driver.TARGET).resolve()
SCRATCH = Path(driver.SCRATCH).resolve()
BASE = "41bb4c9670935496a5ca421d86bc76f7f6485971"
TOOL = "tools/perf-baseline/Cargo.toml"
ALLOWED_SOURCES = (
    "tools/perf-baseline/src/filesystem.rs",
    "tools/perf-baseline/src/filesystem/aligned_zip.rs",
    "tools/perf-baseline/README.md",
)
CASES = (
    "opc_file_eager_open",
    "opc_file_source_open",
    "opc_file_eager_one_part_atomic_save",
    "opc_file_source_one_part_atomic_save",
    "pptx_file_eager_open_selected_slide_lifecycle",
    "pptx_file_source_open_selected_slide_lifecycle",
)
PPTX_CASE = CASES[-1]
UNALIGNED_PPTX_SHA256 = (
    "61b2b99083ca27ebd37955db600955e3f41289b93dba71951983164239eff757"
)
UNALIGNED_PPTX_BYTES = 17_017_139
ALIGNED_PPTX_SHA256 = (
    "6987d620c99cde600b659a88da80ebb16f98bbc2b90633fb113e2a056c13377c"
)
ALIGNED_PPTX_BYTES = 17_018_880
SELECTED_PAYLOAD_BYTES = 522
SELECTED_SEMANTIC_SHA256 = (
    "f5f7db181150c00a4323a48c142721ead73aca3ad7c3b3594e8b1a18a686b257"
)
TAIL_BYTES = 65_536
PLAN_SCHEMA = "litchi.0834.filesystem-plan.v1"
AUDIT_SCHEMA = "litchi.0834.independent-custody-audit.v1"


class AuditError(RuntimeError):
    """The retained packet is missing, stale, or internally contradictory."""


def fail(message: str) -> None:
    raise AuditError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON evidence: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"invalid JSON evidence {path}: {error}")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1 << 20), b""):
                digest.update(block)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def valid_sha(value: Any) -> bool:
    return (isinstance(value, str) and len(value) == 64
            and all(char in "0123456789abcdef" for char in value))


def path_inside(path: Path, root: Path, label: str) -> Path:
    resolved = path.resolve(strict=False)
    require(resolved.is_relative_to(root), f"{label} escaped {root}: {path}")
    return resolved


def packet_path(raw: Any, label: str) -> Path:
    require(isinstance(raw, str) and raw, f"{label} path is missing")
    value = Path(raw)
    return path_inside(value if value.is_absolute() else PACKET / value,
                        PACKET, label)


def root_path(raw: Any, label: str) -> Path:
    require(isinstance(raw, str) and raw and not Path(raw).is_absolute(),
            f"{label} path is invalid")
    return path_inside(ROOT / raw, ROOT, label)


def contains_descriptor(value: Any, expected: dict[str, Any]) -> bool:
    if isinstance(value, dict):
        if all(value.get(key) == expected.get(key)
               for key in ("path", "bytes", "sha256")):
            return True
        return any(contains_descriptor(item, expected) for item in value.values())
    if isinstance(value, list):
        return any(contains_descriptor(item, expected) for item in value)
    return False


def descriptor_value(value: Any, label: str, *, allow_cleanup: bool = False) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label} descriptor is malformed")
    raw_path = value.get("path")
    require(isinstance(raw_path, str) and raw_path, f"{label} descriptor path is missing")
    path = Path(raw_path)
    if not path.is_absolute():
        path = PACKET / path
    require(not path.is_symlink(), f"{label} descriptor is a symlink")
    path = path.resolve(strict=False)
    # Empty command logs are valid evidence (for example, the successful warm
    # diagnostic emits no log output); the hash still binds their exact bytes.
    require(type(value.get("bytes")) is int and value["bytes"] >= 0,
            f"{label} descriptor byte count is invalid")
    require(valid_sha(value.get("sha256")), f"{label} descriptor hash is invalid")
    if path.is_file() and not path.is_symlink():
        require(path.stat().st_size == value["bytes"], f"{label} descriptor bytes changed")
        require(sha256(path) == value["sha256"], f"{label} descriptor hash changed")
    else:
        require(allow_cleanup, f"missing {label} without cleanup allowance")
        witness = PACKET / "cleanup.json"
        require(witness.is_file() and not witness.is_symlink(),
                f"missing cleanup witness for {label}")
        retained = read_json(witness)
        require(retained.get("status") == "pass",
                "cleanup witness is not terminal pass")
        expected = {"path": raw_path, "bytes": value["bytes"],
                    "sha256": value["sha256"]}
        require(contains_descriptor(retained, expected),
                f"cleanup witness does not retain {label}")
    return {"path": raw_path, "bytes": value["bytes"],
            "sha256": value["sha256"]}


def descriptor(path: Path, label: str, *, allow_cleanup: bool = False) -> dict[str, Any]:
    raw = str(path)
    regular = path.is_file() and not path.is_symlink()
    return descriptor_value({"path": raw, "bytes": path.stat().st_size
                             if regular else 1,
                             "sha256": sha256(path) if regular else
                             "0" * 64},
                            label, allow_cleanup=allow_cleanup)


def source_descriptor(path: Path, label: str) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"missing {label}: {path}")
    return {"path": str(path), "bytes": path.stat().st_size,
            "sha256": sha256(path)}


def same_descriptor(left: Any, right: Any, label: str) -> None:
    require(isinstance(left, dict) and isinstance(right, dict),
            f"{label} descriptor is missing")
    for key in ("path", "bytes", "sha256"):
        require(left.get(key) == right.get(key),
                f"{label} {key} differs")


def expected_source_archive(stage: str) -> dict[str, Path]:
    root = PACKET / "sources" / stage
    require(root.is_dir() and not root.is_symlink(),
            f"missing {stage} source archive")
    files: dict[str, Path] = {}
    for path in root.rglob("*"):
        require(not path.is_symlink(),
                f"{stage} source archive contains a symlink")
        if path.is_file() and not path.is_symlink():
            files[str(path.relative_to(root))] = path
    require(set(files) <= set(ALLOWED_SOURCES),
            f"{stage} source archive contains an unapproved path")
    return files


def load_origin() -> dict[str, Any]:
    origin_path = PACKET / "origin.json"
    origin = read_json(origin_path)
    require(origin.get("base") == BASE and
            origin.get("scope") == "Harness repair after retained 0833 failures; iWork excluded",
            "origin identity changed")
    normative = origin.get("normative")
    unrelated = origin.get("unrelated")
    require(isinstance(normative, dict) and normative
            and isinstance(unrelated, dict) and unrelated,
            "origin custody maps are malformed")
    for raw, expected in normative.items():
        require(valid_sha(expected), f"normative hash is malformed: {raw}")
        path = root_path(raw, "normative")
        require(path.is_file() and sha256(path) == expected,
                f"normative input changed: {raw}")
    for raw, expected in unrelated.items():
        require(valid_sha(expected), f"unrelated hash is malformed: {raw}")
        path = root_path(raw, "unrelated")
        require(path.is_file() and sha256(path) == expected,
                f"unrelated input changed: {raw}")
    return {
        "path": str(origin_path.relative_to(PACKET)),
        "bytes": origin_path.stat().st_size,
        "sha256": sha256(origin_path),
        "base": BASE,
        "normative_count": len(normative),
        "unrelated": unrelated,
    }


def load_baseline_source() -> dict[str, str]:
    previous = PACKET.parent / "change-0833" / "prepare.json"
    value = read_json(previous)
    source = value.get("source")
    require(isinstance(source, dict) and source,
            "0833 source inventory is missing")
    require(all(isinstance(raw, str) and valid_sha(digest)
                for raw, digest in source.items()),
            "0833 source inventory is malformed")
    return source


def validate_source_archives(origin: dict[str, Any]) -> dict[str, Any]:
    require(set(ALLOWED_SOURCES) == set(driver.ALLOWED),
            "driver allowed source declaration changed")
    baseline = load_baseline_source()
    require(set(expected_source_archive("diagnostic")) <= set(ALLOWED_SOURCES),
            "diagnostic archive source set changed")
    base_root = PACKET / "base"
    candidate_root = PACKET / "candidate"
    require(base_root.is_dir() and candidate_root.is_dir(),
            "base/candidate archives are missing")
    base_files = {
        str(path.relative_to(base_root)): path
        for path in base_root.rglob("*")
        if path.is_file() and not path.is_symlink()
    }
    candidate_files = {
        str(path.relative_to(candidate_root)): path
        for path in candidate_root.rglob("*")
        if path.is_file() and not path.is_symlink()
    }
    for root_archive, label in ((base_root, "base"),
                                (candidate_root, "candidate")):
        for path in root_archive.rglob("*"):
            require(not path.is_symlink(),
                    f"{label} source archive contains a symlink")
    require(set(base_files) <= set(ALLOWED_SOURCES)
            and set(candidate_files) <= set(ALLOWED_SOURCES),
            "base/candidate archive contains an unapproved source")
    expected_base = {raw for raw in ALLOWED_SOURCES if raw in baseline}
    require(set(base_files) == expected_base,
            "base source archive does not match the 0833 source boundary")
    for raw, path in base_files.items():
        require(sha256(path) == baseline[raw],
                f"base source archive changed: {raw}")
    for raw, path in candidate_files.items():
        require(sha256(path) != baseline.get(raw),
                f"candidate source did not change the baseline: {raw}")

    freeze_path = PACKET / "freeze-diagnostic.json"
    freeze = read_json(freeze_path)
    require(freeze.get("stage") == "diagnostic",
            "diagnostic freeze stage changed")
    require(freeze.get("source") and isinstance(freeze["source"], dict),
            "diagnostic freeze source inventory is missing")
    source = freeze["source"]
    require(set(source) >= set(baseline),
            "diagnostic freeze dropped an input")
    for raw, expected in baseline.items():
        if raw not in ALLOWED_SOURCES:
            require(source.get(raw) == expected,
                    f"diagnostic freeze changed unowned source: {raw}")
    for raw in source:
        require(valid_sha(source[raw]), f"diagnostic freeze hash is malformed: {raw}")
        if raw in ALLOWED_SOURCES:
            archive = PACKET / "sources" / "diagnostic" / raw
            require(archive.is_file() and sha256(archive) == source[raw],
                    f"diagnostic source snapshot changed: {raw}")
        else:
            path = root_path(raw, "diagnostic frozen source")
            require(path.is_file() and sha256(path) == source[raw],
                    f"diagnostic frozen source changed: {raw}")
    diagnostic_archive = expected_source_archive("diagnostic")
    require(set(diagnostic_archive) == (set(source) & set(ALLOWED_SOURCES)),
            "diagnostic source archive is incomplete")
    same_descriptor(freeze.get("driver"),
                    source_descriptor(PACKET / "driver.py", "driver.py"),
                    "diagnostic freeze driver")
    same_descriptor(freeze.get("origin"),
                    source_descriptor(PACKET / "origin.json", "origin.json"),
                    "diagnostic freeze origin")

    repaired_path = PACKET / "freeze-repaired.json"
    repaired = None
    if repaired_path.is_file():
        repaired = read_json(repaired_path)
        require(repaired.get("stage") == "repaired",
                "repaired freeze stage changed")
        repaired_source = repaired.get("source")
        require(isinstance(repaired_source, dict)
                and set(repaired_source) >= set(baseline)
                and set(ALLOWED_SOURCES) <= set(repaired_source),
                "repaired freeze source set is incomplete")
        for raw, expected in baseline.items():
            if raw not in ALLOWED_SOURCES:
                require(repaired_source.get(raw) == expected,
                        f"repaired freeze changed unowned source: {raw}")
        for raw, expected in repaired_source.items():
            require(valid_sha(expected), f"repaired freeze hash is malformed: {raw}")
            if raw in ALLOWED_SOURCES:
                archive = PACKET / "sources" / "repaired" / raw
                require(archive.is_file() and sha256(archive) == expected,
                        f"repaired source snapshot changed: {raw}")
            else:
                path = root_path(raw, "repaired frozen source")
                require(path.is_file() and sha256(path) == expected,
                        f"repaired frozen source changed: {raw}")
        repaired_archive = expected_source_archive("repaired")
        require(set(repaired_archive) == set(ALLOWED_SOURCES),
                "repaired source archive does not contain exactly three sources")
        same_descriptor(repaired.get("driver"),
                        source_descriptor(PACKET / "driver.py", "driver.py"),
                        "repaired freeze driver")
        same_descriptor(repaired.get("origin"),
                        source_descriptor(PACKET / "origin.json", "origin.json"),
                        "repaired freeze origin")
    repaired_v2_path = PACKET / "freeze-repaired-v2.json"
    repaired_v2 = None
    if repaired_v2_path.is_file():
        repaired_v2 = read_json(repaired_v2_path)
        require(repaired_v2.get("stage") == "repaired-v2",
                "repaired-v2 freeze stage changed")
        repaired_v2_source = repaired_v2.get("source")
        require(isinstance(repaired_v2_source, dict)
                and set(repaired_v2_source) >= set(baseline)
                and set(ALLOWED_SOURCES) <= set(repaired_v2_source),
                "repaired-v2 freeze source set is incomplete")
        for raw, expected in baseline.items():
            if raw not in ALLOWED_SOURCES:
                require(repaired_v2_source.get(raw) == expected,
                        f"repaired-v2 freeze changed unowned source: {raw}")
        for raw, expected in repaired_v2_source.items():
            require(valid_sha(expected),
                    f"repaired-v2 freeze hash is malformed: {raw}")
            if raw in ALLOWED_SOURCES:
                archive = PACKET / "sources" / "repaired-v2" / raw
                require(archive.is_file() and sha256(archive) == expected,
                        f"repaired-v2 source snapshot changed: {raw}")
            else:
                path = root_path(raw, "repaired-v2 frozen source")
                require(path.is_file() and sha256(path) == expected,
                        f"repaired-v2 frozen source changed: {raw}")
        repaired_v2_archive = expected_source_archive("repaired-v2")
        require(set(repaired_v2_archive) == set(ALLOWED_SOURCES),
                "repaired-v2 source archive does not contain exactly three sources")
        same_descriptor(repaired_v2.get("driver"),
                        source_descriptor(PACKET / "driver_v2.py", "driver_v2.py"),
                        "repaired-v2 freeze driver")
        same_descriptor(repaired_v2.get("origin"),
                        source_descriptor(PACKET / "origin.json", "origin.json"),
                        "repaired-v2 freeze origin")
    repaired_v3_path = PACKET / "freeze-repaired-v3.json"
    repaired_v3 = None
    if repaired_v3_path.is_file():
        repaired_v3 = read_json(repaired_v3_path)
        require(repaired_v3.get("stage") == "repaired-v3",
                "repaired-v3 freeze stage changed")
        repaired_v3_source = repaired_v3.get("source")
        require(isinstance(repaired_v3_source, dict)
                and set(repaired_v3_source) >= set(baseline)
                and set(ALLOWED_SOURCES) <= set(repaired_v3_source),
                "repaired-v3 freeze source set is incomplete")
        for raw, expected in baseline.items():
            if raw not in ALLOWED_SOURCES:
                require(repaired_v3_source.get(raw) == expected,
                        f"repaired-v3 freeze changed unowned source: {raw}")
        for raw, expected in repaired_v3_source.items():
            require(valid_sha(expected),
                    f"repaired-v3 freeze hash is malformed: {raw}")
            if raw in ALLOWED_SOURCES:
                archive = PACKET / "sources" / "repaired-v3" / raw
                require(archive.is_file() and sha256(archive) == expected,
                        f"repaired-v3 source snapshot changed: {raw}")
            else:
                path = root_path(raw, "repaired-v3 frozen source")
                require(path.is_file() and sha256(path) == expected,
                        f"repaired-v3 frozen source changed: {raw}")
        repaired_v3_archive = expected_source_archive("repaired-v3")
        require(set(repaired_v3_archive) == set(ALLOWED_SOURCES),
                "repaired-v3 source archive does not contain exactly three sources")
        same_descriptor(repaired_v3.get("driver"),
                        source_descriptor(PACKET / "driver_v3.py", "driver_v3.py"),
                        "repaired-v3 freeze driver")
        same_descriptor(repaired_v3.get("origin"),
                        source_descriptor(PACKET / "origin.json", "origin.json"),
                        "repaired-v3 freeze origin")
    return {
        "baseline_source_count": len(baseline),
        "allowed_sources": list(ALLOWED_SOURCES),
        "diagnostic_freeze": source_descriptor(freeze_path, "diagnostic freeze"),
        "diagnostic_source_count": len(diagnostic_archive),
        "base_source_count": len(base_files),
        "candidate_source_count": len(candidate_files),
        "repaired": repaired is not None,
        "repaired_freeze": (source_descriptor(repaired_path, "repaired freeze")
                            if repaired is not None else None),
        "repaired_v2": repaired_v2 is not None,
        "repaired_v2_freeze": (source_descriptor(repaired_v2_path,
                                                   "repaired-v2 freeze")
                                if repaired_v2 is not None else None),
        "repaired_v3": repaired_v3 is not None,
        "repaired_v3_freeze": (source_descriptor(repaired_v3_path,
                                                   "repaired-v3 freeze")
                                if repaired_v3 is not None else None),
    }


def validate_host(origin: dict[str, Any]) -> dict[str, Any]:
    path = PACKET / "host.json"
    value = read_json(path)
    require(value.get("base") == BASE and isinstance(value.get("affinity"), list)
            and 12 in value["affinity"], "host receipt changed")
    return {"path": str(path.relative_to(PACKET)), "bytes": path.stat().st_size,
            "sha256": sha256(path), "base": BASE}


def validate_freeze_for_stage(stage: str) -> dict[str, Any]:
    path = PACKET / f"freeze-{stage}.json"
    value = read_json(path)
    require(value.get("stage") == stage and isinstance(value.get("source"), dict),
            f"{stage} freeze is malformed")
    return {"path": str(path.relative_to(PACKET)), "bytes": path.stat().st_size,
            "sha256": sha256(path), "stage": stage}


def expected_freeze(stage: str) -> Path:
    return PACKET / f"freeze-{stage}.json"


def validate_active_stage() -> dict[str, Any]:
    path = PACKET / "active-stage.json"
    if not path.is_file():
        return {"present": False, "commands": []}
    value = read_json(path)
    require(value.get("stage") == "repaired-v2"
            and value.get("previous_stage") == "repaired",
            "active stage identity changed")
    driver_descriptor = source_descriptor(PACKET / "driver_v2.py", "driver_v2.py")
    same_descriptor(value.get("driver"), driver_descriptor,
                    "active stage driver")
    source_change = value.get("source_change")
    require(isinstance(source_change, dict)
            and source_change.get("path")
            == "tools/perf-baseline/src/filesystem/aligned_zip.rs"
            and source_change.get("before")
            == "aligned.len() % page_size != 0"
            and source_change.get("after")
            == "!aligned.len().is_multiple_of(page_size)"
            and source_change.get("kind")
            == "equivalent divisibility predicate; page size already rejected when zero",
            "active stage amendment description changed")
    old_source = PACKET / "sources" / "repaired" / source_change["path"]
    new_source = PACKET / "sources" / "repaired-v2" / source_change["path"]
    require(old_source.is_file() and not old_source.is_symlink()
            and new_source.is_file() and not new_source.is_symlink(),
            "active stage amendment snapshots are missing")
    old_lines = old_source.read_text(encoding="utf-8").splitlines(keepends=True)
    new_lines = new_source.read_text(encoding="utf-8").splitlines(keepends=True)
    changes = []
    for index in range(max(len(old_lines), len(new_lines))):
        old_line = old_lines[index] if index < len(old_lines) else None
        new_line = new_lines[index] if index < len(new_lines) else None
        if old_line != new_line:
            changes.append((old_line, new_line))
    require(len(changes) == 1,
            "active stage source amendment is not a one-line diff")
    old_line, new_line = changes[0]
    require(isinstance(old_line, str) and isinstance(new_line, str)
            and old_line.count(source_change["before"]) == 1
            and new_line.count(source_change["after"]) == 1
            and old_line.replace(source_change["before"], source_change["after"])
            == new_line,
            "active stage source amendment differs from the pinned predicate")

    full_suite = descriptor_value(value.get("full_suite"),
                                  "active full-suite receipt")
    failed_clippy = descriptor_value(value.get("failed_clippy"),
                                     "active failed-clippy receipt")
    full_suite_actual = descriptor(PACKET / "commands/quality-test/receipt.json",
                                   "active full-suite receipt")
    failed_clippy_actual = descriptor(
        PACKET / "commands/quality-clippy/receipt.json",
        "active failed-clippy receipt")
    same_descriptor(full_suite, full_suite_actual,
                    "active full-suite receipt")
    same_descriptor(failed_clippy, failed_clippy_actual,
                    "active failed-clippy receipt")
    require(value.get("formal_capture")
            == "deferred until final repaired qualification is admitted",
            "active formal-capture status changed")

    original_rows = []
    for gate in ("fmt", "check", "test"):
        original_rows.append(read_command("quality-" + gate, "repaired",
                                          quality_argv(gate), 0))
    original_rows.append(read_command("quality-clippy", "repaired",
                                     quality_argv("clippy"), 101))
    clippy_log = PACKET / "commands/quality-clippy/output.log"
    log_text = clippy_log.read_text(encoding="utf-8")
    error_lines = [line for line in log_text.splitlines()
                   if line.startswith("error: ")]
    require(log_text.count("manual implementation of `.is_multiple_of()`") == 1
            and log_text.count("src/filesystem/aligned_zip.rs:168:49") == 1
            and log_text.count("manual-is-multiple-of") == 1
            and error_lines
            and all(line == "error: manual implementation of `.is_multiple_of()`"
                    or line.startswith("error: could not compile ")
                    for line in error_lines),
            "original clippy failure was not the pinned single helper")
    return {"present": True,
            "path": source_descriptor(path, "active stage"),
            "driver": driver_descriptor,
            "source_change": {"path": source_change["path"],
                               "before": source_change["before"],
                               "after": source_change["after"]},
            "full_suite": full_suite, "failed_clippy": failed_clippy,
            "commands": original_rows}


def validate_follow_stage() -> dict[str, Any]:
    """Validate an additive repaired-v3 amendment when the follow-up exists."""
    path = PACKET / "active-stage-v3.json"
    if not path.is_file():
        return {"present": False, "commands": []}
    value = read_json(path)
    require(value.get("stage") == "repaired-v3"
            and value.get("previous_stage") == "repaired-v2",
            "active repaired-v3 stage identity changed")
    driver_descriptor = source_descriptor(PACKET / "driver_v3.py", "driver_v3.py")
    same_descriptor(value.get("driver"), driver_descriptor,
                    "active repaired-v3 driver")
    source_change = value.get("source_change")
    require(isinstance(source_change, dict)
            and source_change.get("path")
            == "tools/perf-baseline/src/filesystem/aligned_zip.rs",
            "active repaired-v3 source amendment changed")
    require(source_change.get("reason")
            == "Use existing ZIP descriptor validation; compare complete physical local spans including descriptors. Add descriptor framing tests.",
            "active repaired-v3 source amendment reason changed")
    change_path = source_change["path"]
    before_binding = descriptor_value(source_change.get("before"),
                                      "active repaired-v3 source before")
    after_binding = descriptor_value(source_change.get("after"),
                                     "active repaired-v3 source after")
    before_snapshot = source_descriptor(
        PACKET / "sources" / "repaired-v2" / change_path,
        "repaired-v2 source amendment before")
    after_snapshot = source_descriptor(
        PACKET / "sources" / "repaired-v3" / change_path,
        "repaired-v3 source amendment after")
    live_after = source_descriptor(ROOT / change_path,
                                   "live repaired-v3 source amendment")
    same_descriptor(before_binding, before_snapshot,
                    "active repaired-v3 source before")
    same_descriptor(after_binding, live_after,
                    "active repaired-v3 source after")
    require(all(after_snapshot.get(field) == after_binding.get(field)
                for field in ("bytes", "sha256")),
            "repaired-v3 source amendment archive differs from active source")
    require(before_snapshot["sha256"] != after_snapshot["sha256"],
            "repaired-v3 source amendment did not change the helper source")
    for raw in ALLOWED_SOURCES:
        old = PACKET / "sources" / "repaired-v2" / raw
        new = PACKET / "sources" / "repaired-v3" / raw
        require(old.is_file() and new.is_file(),
                f"repaired-v3 comparison source is missing: {raw}")
        if raw != change_path:
            require(old.read_bytes() == new.read_bytes(),
                    f"repaired-v3 changed an unowned source: {raw}")

    failed_qualification = PACKET / "qualification-v2.json"
    failed_descriptor = descriptor(failed_qualification,
                                   "failed v2 qualification")
    same_descriptor(value.get("failed_qualification"), failed_descriptor,
                    "active repaired-v3 failed qualification")
    full_suite = descriptor(PACKET / "commands/quality-test/receipt.json",
                            "active repaired-v3 full suite")
    same_descriptor(value.get("full_suite"), full_suite,
                    "active repaired-v3 full suite")
    previous_amendment = descriptor(PACKET / "active-stage.json",
                                    "active repaired-v2 amendment")
    same_descriptor(value.get("previous_amendment"), previous_amendment,
                    "active repaired-v3 previous amendment")
    require(value.get("formal_capture")
            == "deferred to next batch after the harness repair is committed",
            "active repaired-v3 formal-capture status changed")
    return {"present": True, "path": source_descriptor(path, "active v3 stage"),
            "driver": driver_descriptor,
            "source_change": {"path": change_path,
                               "before": before_binding,
                               "after": after_binding,
                               "reason": source_change.get("reason")},
            "full_suite": full_suite,
            "failed_qualification": failed_descriptor,
            "previous_amendment": previous_amendment,
            "commands": []}


def read_command(label: str, stage: str, expected_argv: list[str],
                 expected_exit: int) -> dict[str, Any]:
    root = PACKET / "commands" / label
    require(root.is_dir() and not root.is_symlink(),
            f"missing command directory: {label}")
    started_path = root / "started.json"
    receipt_path = root / "receipt.json"
    log_path = root / "output.log"
    started = read_json(started_path)
    receipt = read_json(receipt_path)
    require(started.get("argv") == expected_argv
            and receipt.get("argv") == expected_argv,
            f"{label} argv changed")
    require(started.get("cwd") == str(ROOT),
            f"{label} cwd changed")
    require(started.get("started_unix") == receipt.get("started_unix")
            and type(started.get("started_unix")) in (int, float)
            and not isinstance(started.get("started_unix"), bool)
            and math.isfinite(float(started["started_unix"]))
            and type(receipt.get("finished_unix")) in (int, float)
            and not isinstance(receipt.get("finished_unix"), bool)
            and math.isfinite(float(receipt["finished_unix"]))
            and receipt["finished_unix"] >= receipt["started_unix"],
            f"{label} chronology fields changed")
    require(receipt.get("exit_code") == expected_exit
            and receipt.get("error") is None,
            f"{label} exit/error state changed")
    freeze = descriptor_value(started.get("freeze"), f"{label} start freeze")
    terminal_freeze = descriptor_value(receipt.get("freeze"), f"{label} receipt freeze")
    expected = descriptor_value(
        {"path": str(expected_freeze(stage)),
         "bytes": expected_freeze(stage).stat().st_size,
         "sha256": sha256(expected_freeze(stage))},
        f"{label} expected freeze")
    same_descriptor(freeze, expected, f"{label} start freeze")
    same_descriptor(terminal_freeze, expected, f"{label} receipt freeze")
    log = descriptor_value(receipt.get("log"), f"{label} log")
    same_descriptor(log, descriptor(log_path, f"{label} output log"),
                    f"{label} log")
    return {
        "name": label,
        # Driver descriptors use absolute packet paths.  Preserve that exact
        # spelling here so receipt bindings can be compared byte-for-byte;
        # relative paths are used only in human-facing labels below.
        "started": descriptor(started_path, f"{label} started"),
        "receipt": descriptor(receipt_path, f"{label} receipt"),
        "log": descriptor(log_path, f"{label} output log"),
        "argv": expected_argv,
        "exit_code": expected_exit,
        "started_unix": started["started_unix"],
        "finished_unix": receipt["finished_unix"],
        "stage": stage,
    }


def validate_serial(rows: list[dict[str, Any]]) -> None:
    ordered = sorted(rows, key=lambda row: (row["started_unix"], row["name"]))
    for previous, current in zip(ordered, ordered[1:]):
        require(previous["finished_unix"] <= current["started_unix"],
                f"command intervals overlap: {previous['name']} and {current['name']}")


def validate_command_inventory(rows: list[dict[str, Any]]) -> None:
    root = PACKET / "commands"
    require(root.is_dir() and not root.is_symlink(),
            "command receipt root is missing")
    actual = set()
    for path in root.iterdir():
        require(path.is_dir() and not path.is_symlink(),
                f"unexpected command receipt entry: {path.name}")
        actual.add(path.name)
    expected = [row["name"] for row in rows]
    require(len(expected) == len(set(expected)) and actual == set(expected),
            "command receipt inventory changed")


def diagnostic_argv(binary: str, state: str, report: Path) -> list[str]:
    return ["taskset", "-c", "12", binary, "--case", PPTX_CASE,
            "--samples", "1", "--warmup", "0", "--filesystem-cache", state,
            "--filesystem-root", str(SCRATCH), "--json", str(report)]


def validate_build_diagnostic() -> tuple[dict[str, Any], dict[str, Any]]:
    build_path = PACKET / "build-diagnostic.json"
    value = read_json(build_path)
    binary = descriptor_value(value.get("binary"), "diagnostic binary",
                              allow_cleanup=True)
    binary_path = Path(binary["path"]).resolve(strict=False)
    require(binary_path.name == "litchi-perf-baseline"
            and binary_path.parent.name == "diagnostic"
            and binary_path.is_relative_to(TARGET),
            "diagnostic binary path changed")
    receipt = descriptor_value(value.get("receipt"), "diagnostic build receipt")
    command = read_command(
        "build-diagnostic", "diagnostic",
        ["cargo", "build", "--offline", "--locked", "--release",
         "--manifest-path", TOOL, "--bin", "litchi-perf-baseline"], 0)
    same_descriptor(receipt, command["receipt"], "diagnostic build receipt")
    return ({"binary": binary, "receipt": receipt},
            command)


def check_report_descriptor(value: Any, path: Path, label: str) -> dict[str, Any]:
    actual = descriptor(path, label)
    same_descriptor(value, actual, label)
    return actual


def validate_warm_report(path: Path, binary: dict[str, Any]) -> dict[str, Any]:
    report = read_json(path)
    require(report.get("schema_version") == 1, "warm diagnostic report schema changed")
    tool = report.get("tool")
    require(isinstance(tool, dict) and tool.get("name") == "litchi-perf-baseline"
            and tool.get("binary") == "litchi-perf-baseline"
            and tool.get("instrumentation") == "none",
            "warm diagnostic tool identity changed")
    identity = report.get("binary_identity")
    require(isinstance(identity, dict)
            and identity.get("path") == binary["path"]
            and identity.get("binary_sha256") == binary["sha256"]
            and identity.get("binary_bytes") == binary["bytes"]
            and identity.get("executable") is True
            and identity.get("profile") == "release",
            "warm diagnostic binary binding changed")
    environment = report.get("environment")
    if isinstance(environment, dict) and environment.get("git_revision") is not None:
        require(environment["git_revision"] == BASE,
                "warm diagnostic build revision changed")
    configuration = report.get("configuration")
    require(isinstance(configuration, dict)
            and configuration.get("samples_per_case") == 1
            and configuration.get("warmup_iterations_per_case") == 0
            and configuration.get("filesystem_cache_states") == ["warm"]
            and configuration.get("cases") == [PPTX_CASE],
            "warm diagnostic configuration changed")
    results = report.get("results")
    require(isinstance(results, list) and len(results) == 1,
            "warm diagnostic result count changed")
    result = results[0]
    elapsed = result.get("elapsed_ns")
    evidence = report.get("filesystem_evidence")
    require(isinstance(evidence, list) and len(evidence) == 1
            and isinstance(evidence[0], dict)
            and isinstance(evidence[0].get("samples"), list)
            and len(evidence[0]["samples"]) == 1,
            "warm diagnostic filesystem evidence changed")
    evidence_sample = evidence[0]["samples"][0]
    replay = evidence_sample.get("pptx_source_replay")
    require(result.get("case") == PPTX_CASE and result.get("cache_state") == "warm"
            and isinstance(elapsed, dict)
            and elapsed.get("samples") and len(elapsed["samples"]) == 1
            and elapsed.get("sample_order") == [0]
            and evidence[0].get("case") == PPTX_CASE
            and evidence_sample.get("cache_state") == "warm"
            and evidence_sample.get("elapsed_ns") == elapsed["samples"][0]
            and isinstance(replay, dict),
            "warm diagnostic result identity changed")
    require(replay.get("source_bytes") == UNALIGNED_PPTX_BYTES
            and replay.get("source_sha256") == UNALIGNED_PPTX_SHA256
            and replay.get("selected_position") == 100
            and replay.get("slide_count") == 200
            and replay.get("read_calls") == 882
            and replay.get("read_bytes") == 104171
            and replay.get("selected_slide_payload_read_bytes") == SELECTED_PAYLOAD_BYTES
            and replay.get("selected_slide_payload_fully_covered") is True
            and replay.get("unselected_slide_payload_read_bytes") == 0
            and replay.get("media_payload_read_bytes") == 0
            and replay.get("semantic_sha256") == SELECTED_SEMANTIC_SHA256
            and replay.get("classification")
            == "selected-slide-only:target-slide-no-unselected-or-media-overlap",
            "warm diagnostic source replay changed")
    return {"path": str(path.relative_to(PACKET)), "bytes": path.stat().st_size,
            "sha256": sha256(path), "source_sha256": UNALIGNED_PPTX_SHA256,
            "source_bytes": UNALIGNED_PPTX_BYTES}


def parse_nested_diagnostic(log: Path) -> dict[str, Any]:
    text = log.read_text(encoding="utf-8")
    line = text.strip()
    require(line.startswith("Error: "), "cold diagnostic log prefix changed")
    try:
        outer = json.loads(line.removeprefix("Error: "))
        inner = json.loads(outer.split(": Error: ", 1)[1].strip())
        payload = inner.split("; diagnostic=", 1)[1]
        return json.loads(payload)
    except (IndexError, UnicodeError, json.JSONDecodeError, TypeError) as error:
        fail(f"cold diagnostic log cannot be decoded: {error}")


def checked_ranges(value: Any, label: str, source_bytes: int) -> list[dict[str, int]]:
    require(isinstance(value, list), f"{label} ranges are missing")
    result = []
    for index, row in enumerate(value):
        require(isinstance(row, dict)
                and type(row.get("start")) is int
                and type(row.get("end")) is int
                and 0 <= row["start"] < row["end"] <= source_bytes,
                f"{label}[{index}] range is invalid")
        result.append({"start": row["start"], "end": row["end"]})
    ordered = sorted(result, key=lambda item: (item["start"], item["end"]))
    for previous, current in zip(ordered, ordered[1:]):
        require(previous["end"] <= current["start"],
                f"{label} ranges overlap")
    return result


def overlap(read: dict[str, int], ranges: Iterable[dict[str, int]]) -> int:
    start = read["offset"]
    end = start + read["returned_length"]
    return sum(max(0, min(end, row["end"]) - max(start, row["start"]))
               for row in ranges)


def merged_ranges(values: Iterable[dict[str, int]]) -> list[dict[str, int]]:
    ordered = sorted(values, key=lambda item: (item["start"], item["end"]))
    result: list[dict[str, int]] = []
    for row in ordered:
        if result and row["start"] <= result[-1]["end"]:
            result[-1]["end"] = max(result[-1]["end"], row["end"])
        else:
            result.append(dict(row))
    return result


def phase_counters(reads: list[dict[str, int]],
                   groups: dict[str, list[dict[str, int]]]) -> dict[str, int]:
    result = {"read_calls": len(reads),
              "read_bytes": sum(row["returned_length"] for row in reads)}
    for name, ranges in groups.items():
        values = [overlap(row, ranges) for row in reads]
        result[f"{name}_payload_read_calls"] = sum(value > 0 for value in values)
        result[f"{name}_payload_read_bytes"] = sum(values)
    return result


def phase_coverage(reads: list[dict[str, int]],
                   groups: dict[str, list[dict[str, int]]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for name, ranges in groups.items():
        # The diagnostic calls this one range group "selected" in its
        # observed-range field, while its counters use "selected_slide".
        coverage_name = "selected" if name == "selected_slide" else name
        observed = []
        for read in reads:
            start = read["offset"]
            end = start + read["returned_length"]
            for item in ranges:
                left = max(start, item["start"])
                right = min(end, item["end"])
                if left < right:
                    observed.append({"start": left, "end": right})
        merged = merged_ranges(observed)
        if name != "unselected_slide":
            result[f"{coverage_name}_observed_ranges"] = merged
        result[f"{name}_payload_covered_bytes"] = sum(
            item["end"] - item["start"] for item in merged)
        result[f"{name}_payload_total_bytes"] = sum(
            item["end"] - item["start"] for item in ranges)
        full_count = sum(
            1 for item in ranges
            if any(observed["start"] <= item["start"]
                   and observed["end"] >= item["end"] for observed in merged))
        if name == "slide":
            result["slide_payload_ranges_fully_covered"] = full_count
        elif name == "selected_slide":
            result["selected_slide_payload_fully_covered"] = (
                full_count == len(ranges))
    return result


def validate_raw_diagnostic(log: Path) -> tuple[dict[str, Any], dict[str, Any]]:
    decoded = parse_nested_diagnostic(log)
    require(decoded.get("operation") == PPTX_CASE
            and decoded.get("child_mode") == "verified-prime"
            and decoded.get("classification") == "classification-failed"
            and decoded.get("source_bytes") == ALIGNED_PPTX_BYTES
            and decoded.get("source_sha256") == ALIGNED_PPTX_SHA256,
            "cold diagnostic envelope changed")
    source_bytes = decoded["source_bytes"]
    payload_ranges = decoded.get("payload_ranges")
    require(isinstance(payload_ranges, dict), "cold payload ranges are missing")
    selected = payload_ranges.get("selected_slide")
    if isinstance(selected, dict):
        selected = [selected]
    groups = {
        "slide": checked_ranges(payload_ranges.get("slides"),
                                "slide payload", source_bytes),
        "selected_slide": checked_ranges(selected,
                                         "selected payload", source_bytes),
        "unselected_slide": checked_ranges(payload_ranges.get("unselected_slides"),
                                           "unselected payload", source_bytes),
        "media": checked_ranges(payload_ranges.get("media"),
                                "media payload", source_bytes),
    }
    require({name: len(values) for name, values in groups.items()} == {
        "slide": 200, "selected_slide": 1,
        "unselected_slide": 199, "media": 8},
            "cold payload range counts changed")
    require(decoded.get("payload_counts") == {
        "slides": 200, "selected_slide": 1,
        "unselected_slides": 199, "media": 8},
            "cold payload count oracle changed")
    require(groups["selected_slide"][0] in groups["slide"]
            and not any(overlap({"offset": row["start"],
                                  "returned_length": row["end"] - row["start"]},
                                 groups["unselected_slide"])
                        for row in groups["selected_slide"]),
            "selected/unselected payload ranges overlap")
    raw_reads = decoded.get("raw_reads")
    require(isinstance(raw_reads, list) and raw_reads,
            "cold raw-read vector is missing")
    reads: list[dict[str, int]] = []
    for index, row in enumerate(raw_reads):
        require(isinstance(row, dict)
                and type(row.get("offset")) is int
                and type(row.get("requested_length")) is int
                and type(row.get("returned_length")) is int
                and row["offset"] >= 0
                and 0 <= row["returned_length"] <= row["requested_length"]
                and row["offset"] + row["returned_length"] <= source_bytes,
                f"cold raw read {index} is invalid")
        reads.append({"offset": row["offset"],
                      "requested_length": row["requested_length"],
                      "returned_length": row["returned_length"]})
    boundary = decoded.get("open_read_count")
    require(type(boundary) is int and 0 <= boundary <= len(reads),
            "cold open-read boundary changed")
    counters = decoded.get("counters")
    require(isinstance(counters, dict), "cold counters are missing")
    all_counts = phase_counters(reads, groups)
    require(counters == all_counts, "cold counters do not recompute from raw reads")
    phases = {
        "all": reads,
        "open": reads[:boundary],
        "query": reads[boundary:],
    }
    phase_counts = {name: phase_counters(items, groups)
                    for name, items in phases.items()}
    coverage = decoded.get("coverage")
    require(isinstance(coverage, dict), "cold coverage is missing")
    all_coverage = phase_coverage(reads, groups)
    for name, expected in all_coverage.items():
        if name.endswith("_observed_ranges"):
            observed = checked_ranges(coverage.get(name), name, source_bytes)
            # The diagnostic preserves read discovery order, while this
            # independent recomputation merges intervals in offset order.
            require(sorted(observed, key=lambda item: (item["start"], item["end"]))
                    == sorted(expected,
                              key=lambda item: (item["start"], item["end"])),
                    f"cold {name} does not recompute")
        else:
            require(coverage.get(name) == expected,
                    f"cold {name} does not recompute")
    tail = {"offset": source_bytes - TAIL_BYTES,
            "requested_length": TAIL_BYTES,
            "returned_length": TAIL_BYTES}
    tail_indices = [index for index, row in enumerate(reads) if row == tail]
    require(len(tail_indices) == 1 and tail_indices[0] < boundary,
            "cold exact aligned tail read changed")
    tail_index = tail_indices[0]
    open_counts = phase_counts["open"]
    for name, ranges in groups.items():
        require(open_counts[f"{name}_payload_read_bytes"] == overlap(tail, ranges),
                f"cold open {name} overlap is not explained by exact tail")
    require(phase_counts["query"]["selected_slide_payload_read_bytes"]
            == SELECTED_PAYLOAD_BYTES
            and phase_counts["query"]["unselected_slide_payload_read_bytes"] == 0
            and phase_counts["query"]["media_payload_read_bytes"] == 0,
            "cold semantic query overlap changed")
    derived = {
        "source_bytes": source_bytes,
        "source_sha256": decoded["source_sha256"],
        "open_read_count": boundary,
        "tail_read_index": tail_index,
        "tail": tail,
        "counters": all_counts,
        "phases": phase_counts,
        "coverage": all_coverage,
    }
    return decoded, derived


def validate_diagnostic_analysis(decoded: dict[str, Any],
                                 derived: dict[str, Any],
                                 cold_log: Path) -> dict[str, Any]:
    decoded_path = PACKET / "diagnostic-decoded.json"
    analysis_path = PACKET / "diagnostic-analysis.json"
    retained_decoded = read_json(decoded_path)
    require(retained_decoded == decoded,
            "diagnostic-decoded.json differs from raw cold diagnostic")
    analysis = read_json(analysis_path)
    require(analysis.get("status") == "pass"
            and analysis.get("child_mode") == "verified-prime"
            and analysis.get("source_bytes") == ALIGNED_PPTX_BYTES
            and analysis.get("source_sha256") == ALIGNED_PPTX_SHA256
            and analysis.get("open_read_count") == derived["open_read_count"]
            and analysis.get("tail_read_index") == derived["tail_read_index"]
            and analysis.get("tail") == derived["tail"]
            and analysis.get("phases") == derived["phases"]
            and analysis.get("claim")
            == "Failed untimed replay only; no timed cold performance claim.",
            "diagnostic analysis changed")
    require(analysis.get("log_sha256") == sha256(cold_log),
            "diagnostic analysis log binding changed")
    return {
        "decoded": source_descriptor(decoded_path, "diagnostic-decoded"),
        "analysis": source_descriptor(analysis_path, "diagnostic-analysis"),
        "analyzer": source_descriptor(PACKET / "analyze_diagnostic.py",
                                      "diagnostic analyzer"),
        "raw_recomputed": {
            "source_bytes": derived["source_bytes"],
            "source_sha256": derived["source_sha256"],
            "open_read_count": derived["open_read_count"],
            "tail_read_index": derived["tail_read_index"],
            "counters": derived["counters"],
            "phases": derived["phases"],
        },
    }


def validate_diagnostic_results(build: dict[str, Any],
                                commands: dict[str, dict[str, Any]]) -> dict[str, Any]:
    path = PACKET / "diagnostic-results.json"
    value = read_json(path)
    require(value.get("performance_claim") == "none"
            and isinstance(value.get("rows"), list)
            and len(value["rows"]) == 2,
            "diagnostic result summary changed")
    by_state = {row.get("state"): row for row in value["rows"]}
    require(set(by_state) == {"warm", "cold-verified"},
            "diagnostic result states changed")
    warm = by_state["warm"]
    cold = by_state["cold-verified"]
    require(warm.get("case") == PPTX_CASE and warm.get("exit_code") == 0
            and cold.get("case") == PPTX_CASE and cold.get("exit_code") == 1
            and cold.get("report") is None,
            "diagnostic result exit boundary changed")
    warm_report_path = packet_path(warm["report"]["path"], "warm diagnostic report")
    cold_receipt_path = packet_path(cold["receipt"]["path"], "cold diagnostic receipt")
    require(warm_report_path == PACKET / "diagnostic-pptx-warm.json",
            "warm diagnostic report path changed")
    require(cold_receipt_path == PACKET / "commands/diagnostic-pptx-cold-verified/receipt.json",
            "cold diagnostic receipt path changed")
    same_descriptor(warm["receipt"],
                    commands["diagnostic-pptx-warm"]["receipt"],
                    "warm diagnostic summary receipt")
    same_descriptor(cold["receipt"],
                    commands["diagnostic-pptx-cold-verified"]["receipt"],
                    "cold diagnostic summary receipt")
    same_descriptor(warm["report"],
                    descriptor(warm_report_path, "warm diagnostic report"),
                    "warm diagnostic summary report")
    require(not (PACKET / "diagnostic-pptx-cold-verified.json").exists()
            and not (PACKET / "diagnostic-pptx-cold-verified.json").is_symlink(),
            "failed cold primer unexpectedly retained a report")
    return {"path": str(path.relative_to(PACKET)), "bytes": path.stat().st_size,
            "sha256": sha256(path), "warm_report": warm["report"],
            "cold_report_retained": False}


def validate_diagnostic(build_info: dict[str, Any],
                        build_command: dict[str, Any]) -> dict[str, Any]:
    warm_report = PACKET / "diagnostic-pptx-warm.json"
    warm_command = read_command(
        "diagnostic-pptx-warm", "diagnostic",
        diagnostic_argv(build_info["binary"]["path"], "warm", warm_report), 0)
    cold_report = PACKET / "diagnostic-pptx-cold-verified.json"
    cold_command = read_command(
        "diagnostic-pptx-cold-verified", "diagnostic",
        diagnostic_argv(build_info["binary"]["path"], "cold-verified", cold_report), 1)
    cold_log = PACKET / "commands/diagnostic-pptx-cold-verified/output.log"
    require("PPTX source replay violated pptx_file_source_open_selected_slide_lifecycle "
            "payload-range classification" in cold_log.read_text(encoding="utf-8"),
            "cold primer failure text changed")
    decoded, derived = validate_raw_diagnostic(cold_log)
    analysis = validate_diagnostic_analysis(decoded, derived, cold_log)
    results = validate_diagnostic_results(
        build_info, {"diagnostic-pptx-warm": warm_command,
                      "diagnostic-pptx-cold-verified": cold_command})
    warm_summary = validate_warm_report(warm_report, build_info["binary"])
    return {
        "build": build_command["receipt"],
        "warm_command": warm_command["receipt"],
        "cold_command": cold_command["receipt"],
        "warm_report": warm_summary,
        "results": results,
        "analysis": analysis,
        "primer": {"warm_exit": 0, "cold_exit": 1,
                   "cold_report_retained": False,
                   "classification": decoded["classification"]},
    }


def quality_argv(gate: str) -> list[str]:
    prefix = ["--offline", "--locked", "--manifest-path", TOOL]
    return {
        "fmt": ["cargo", "fmt", "--manifest-path", TOOL, "--", "--check"],
        "check": ["cargo", "check", *prefix, "--all-features", "--all-targets"],
        "test": ["cargo", "test", *prefix, "--lib", "--features",
                 "allocator-metrics,ordinary-save-process-metrics",
                 "--", "--test-threads=2"],
        "clippy": ["cargo", "clippy", *prefix, "--all-features", "--all-targets",
                   "--", "-D", "warnings"],
        "doc": ["cargo", "doc", *prefix, "--all-features", "--no-deps"],
        "boundaries": ["python3", "-B", "tools/check_crate_boundaries.py"],
    }[gate]


def quality_v2_argv(gate: str) -> list[str]:
    if gate == "test":
        prefix = ["--offline", "--locked", "--manifest-path", TOOL]
        return ["cargo", "test", *prefix, "--lib", "--features",
                "allocator-metrics,ordinary-save-process-metrics",
                "filesystem::aligned_zip::tests", "--", "--test-threads=2"]
    return quality_argv(gate)


def quality_v3_argv(gate: str, test_scope: str) -> list[str]:
    if gate == "test":
        prefix = ["--offline", "--locked", "--manifest-path", TOOL]
        return ["cargo", "test", *prefix, "--lib", "--features",
                "allocator-metrics,ordinary-save-process-metrics", test_scope,
                "--", "--test-threads=2"]
    return quality_argv(gate)


def build_argv() -> list[str]:
    return ["cargo", "build", "--offline", "--locked", "--release",
            "--manifest-path", TOOL, "--bin", "litchi-perf-baseline"]


def validate_build_stage(stage: str) -> tuple[dict[str, Any], dict[str, Any]]:
    """Validate a retained build and bind it to its exact build receipt."""
    build_path = PACKET / f"build-{stage}.json"
    value = read_json(build_path)
    binary = descriptor_value(value.get("binary"), f"{stage} binary",
                              allow_cleanup=True)
    binary_path = Path(binary["path"]).resolve(strict=False)
    require(binary_path.name == "litchi-perf-baseline"
            and binary_path.parent.name == stage
            and binary_path.is_relative_to(TARGET),
            f"{stage} binary path changed")
    receipt = descriptor_value(value.get("receipt"), f"{stage} build receipt")
    command = read_command(f"build-{stage}", stage, build_argv(), 0)
    same_descriptor(receipt, command["receipt"], f"{stage} build receipt")
    return ({"binary": binary, "receipt": receipt,
             "path": str(build_path.relative_to(PACKET)),
             "bytes": build_path.stat().st_size,
             "sha256": sha256(build_path)}, command)


def workload_argv(binary: str, case: str, state: str, samples: int,
                  warmup: int, report: Path) -> list[str]:
    return ["taskset", "-c", "12", binary, "--case", case,
            "--samples", str(samples), "--warmup", str(warmup),
            "--filesystem-cache", state, "--filesystem-root", str(SCRATCH),
            "--json", str(report)]


def validate_workload_row(row: Any, label: str, stage: str, binary: dict[str, Any],
                          case: str, state: str, samples: int,
                          warmup: int) -> dict[str, Any]:
    """Check custody and result/report bindings for one workload."""
    require(isinstance(row, dict)
            and row.get("case") == case
            and row.get("state") == state
            and row.get("exit_code") == 0,
            f"{label} workload row changed")
    report = PACKET / f"{label}.json"
    command = read_command(label, stage,
                           workload_argv(binary["path"], case, state, samples,
                                         warmup, report), 0)
    require(report.is_file() and not report.is_symlink(),
            f"{label} report is missing")
    same_descriptor(row.get("report"), descriptor(report, f"{label} report"),
                    f"{label} report")
    same_descriptor(row.get("receipt"), command["receipt"],
                    f"{label} receipt")
    return command


def validate_failed_workload_row(row: Any, label: str, stage: str,
                                 binary: dict[str, Any], case: str,
                                 state: str) -> dict[str, Any]:
    """Validate an intentionally retained terminal workload failure."""
    require(isinstance(row, dict)
            and row.get("case") == case
            and row.get("state") == state
            and row.get("exit_code") == 1
            and row.get("report") is None,
            f"{label} failed workload row changed")
    report = PACKET / f"{label}.json"
    require(not report.exists() and not report.is_symlink(),
            f"{label} unexpectedly retained a failed report")
    command = read_command(label, stage,
                           workload_argv(binary["path"], case, state, 1, 0,
                                         report), 1)
    log = PACKET / "commands" / label / "output.log"
    require(log.read_text(encoding="utf-8").strip()
            == 'Error: ProofError("cold ZIP proof does not accept '
               'data-descriptor member framing")',
            f"{label} terminal proof error changed")
    same_descriptor(row.get("receipt"), command["receipt"],
                    f"{label} receipt")
    return command


def count_qualification_report(path: Path, label: str,
                               expected_states: int = 2) -> int:
    """Count retained warm/cold evidence samples without owning report schema."""
    report = read_json(path)
    results = report.get("results")
    evidence = report.get("filesystem_evidence")
    require(isinstance(results, list) and len(results) == expected_states
            and isinstance(evidence, list) and len(evidence) * 2 == expected_states,
            f"{label} retained report state count changed")
    count = 0
    for item in evidence:
        samples = item.get("samples") if isinstance(item, dict) else None
        require(isinstance(samples, list) and len(samples) == 2
                and sorted(sample.get("cache_state") for sample in samples)
                == ["cold-verified", "warm"],
                f"{label} retained report sample count changed")
        count += len(samples)
    require(count == expected_states,
            f"{label} retained report samples changed")
    return count


def validate_optional_pptx_reader_tests() -> dict[str, Any]:
    """Bind the retained mutation count to the v2 PPTX qualification report."""
    receipt_path = PACKET / "pptx-reader-tests.json"
    script_path = PACKET / "test_pptx_reader.py"
    if not receipt_path.exists() and not script_path.exists():
        return {"present": False}
    require(receipt_path.is_file() and not receipt_path.is_symlink()
            and script_path.is_file() and not script_path.is_symlink(),
            "PPTX reader mutation evidence is incomplete")
    report_path = PACKET / "qualification-v2-05.json"
    value = read_json(receipt_path)
    mutations = value.get("mutations")
    names = [
        "missing-aligned-proof", "wrong-base-hash", "wrong-padding",
        "wrong-eocd", "wrong-boundary", "missing-tail", "duplicate-tail",
        "hidden-raw-overlap", "overstated-return", "wrong-selected-range",
        "wrong-semantics-in-both-states",
    ]
    require(value.get("status") == "pass"
            and value.get("report_sha256") == sha256(report_path)
            and isinstance(mutations, list)
            and [item.get("name") for item in mutations] == names
            and all(item.get("rejected") is True for item in mutations),
            "PPTX reader mutation evidence changed")
    return {"present": True,
            "script": source_descriptor(script_path, "PPTX reader test"),
            "receipt": source_descriptor(receipt_path,
                                          "PPTX reader test receipt"),
            "report": source_descriptor(report_path,
                                         "PPTX reader mutation report"),
            "mutations": len(mutations)}


def validate_optional_repaired(active: dict[str, Any]) -> dict[str, Any]:
    stage = "repaired-v2"
    repaired = PACKET / "freeze-repaired-v2.json"
    build_path = PACKET / "build-repaired-v2.json"
    quality = PACKET / "quality-v2.json"
    qualification = PACKET / "qualification-v2.json"
    artifacts = (repaired, build_path, quality, qualification)
    if not any(path.exists() for path in artifacts):
        return {"present": False, "commands": []}
    require(all(path.is_file() and not path.is_symlink() for path in artifacts),
            "repaired-v2 stage is incomplete")
    require(active.get("present") is True,
            "repaired-v2 artifacts exist without active-stage amendment")
    freeze = validate_freeze_for_stage(stage)
    build_info, build_command = validate_build_stage(stage)

    quality_value = read_json(quality)
    gates = ["fmt", "check", "test", "clippy", "doc", "boundaries"]
    require(quality_value == {
        "status": "pass", "gates": gates, "reused": False,
        "test_scope": "filesystem::aligned_zip::tests",
        "full_suite": descriptor(PACKET / "commands/quality-test/receipt.json",
                                  "v2 full-suite receipt"),
        "amendment": descriptor(PACKET / "active-stage.json",
                                 "v2 amendment")},
            "quality-v2 receipt changed")
    quality_rows = [read_command("quality-v2-" + gate, stage,
                                 quality_v2_argv(gate), 0)
                    for gate in gates]

    qvalue = read_json(qualification)
    require(qvalue.get("status") == "failed"
            and isinstance(qvalue.get("rows"), list)
            and len(qvalue["rows"]) == 7,
            "qualification-v2 receipt changed")
    qlabels = [f"qualification-v2-{index:02}" for index in range(6)]
    qlabels.append("qualification-v2-opc-pair")
    qrows = []
    retained_reports = 0
    retained_samples = 0
    terminal_errors = 0
    for index, label in enumerate(qlabels):
        case = CASES[index] if index < 6 else ",".join(CASES[2:4])
        if index in (2, 3, 6):
            qrows.append(validate_failed_workload_row(
                qvalue["rows"][index], label, stage, build_info["binary"],
                case, "warm,cold-verified"))
            terminal_errors += 1
        else:
            qrows.append(validate_workload_row(
                qvalue["rows"][index], label, stage, build_info["binary"],
                case, "warm,cold-verified", 1, 0))
            retained_reports += 1
            retained_samples += count_qualification_report(
                PACKET / f"{label}.json", label)
    require(retained_reports == 4 and retained_samples == 8
            and terminal_errors == 3,
            "qualification-v2 retained counts changed")
    validate_serial([build_command] + quality_rows + qrows)
    return {"present": True, "stage": stage, "admitted": False,
            "freeze": freeze,
            "build": build_info, "build_command": build_command,
            "quality": {"path": str(quality.relative_to(PACKET)),
                        "bytes": quality.stat().st_size,
                        "sha256": sha256(quality),
                        "commands": quality_rows},
            "qualification": {"path": str(qualification.relative_to(PACKET)),
                              "bytes": qualification.stat().st_size,
                              "sha256": sha256(qualification),
                              "commands": qrows,
                              "status": "failed",
                              "retained_reports": retained_reports,
                              "retained_samples": retained_samples,
                              "terminal_errors": terminal_errors},
            "commands": [build_command] + quality_rows + qrows}


def validate_optional_repaired_v3(follow: dict[str, Any],
                                  repaired_v2: dict[str, Any]) -> dict[str, Any]:
    stage = "repaired-v3"
    freeze_path = PACKET / "freeze-repaired-v3.json"
    build_path = PACKET / "build-repaired-v3.json"
    quality_path = PACKET / "quality-v3.json"
    qualification_path = PACKET / "qualification-v3.json"
    artifacts = (freeze_path, build_path, quality_path, qualification_path)
    if not any(path.exists() for path in artifacts):
        require(not follow.get("present"),
                "active repaired-v3 stage is incomplete")
        return {"present": False, "commands": []}
    require(follow.get("present") is True
            and repaired_v2.get("present") is True,
            "repaired-v3 artifacts lack v2 follow-up custody")
    require(all(path.is_file() and not path.is_symlink() for path in artifacts),
            "repaired-v3 stage is incomplete")
    freeze = validate_freeze_for_stage(stage)
    build_info, build_command = validate_build_stage(stage)

    quality_value = read_json(quality_path)
    gates = ["fmt", "check", "test", "clippy", "doc", "boundaries"]
    test_scope = quality_value.get("test_scope", "filesystem::aligned_zip::tests")
    require(quality_value.get("status") == "pass"
            and quality_value.get("gates") == gates
            and quality_value.get("reused") is False
            and isinstance(test_scope, str) and test_scope,
            "quality-v3 receipt changed")
    full_suite_descriptor = descriptor(
        PACKET / "commands/quality-test/receipt.json",
        "v3 full-suite receipt")
    full_bindings = [(key, item) for key, item in quality_value.items()
                     if "full" in key and isinstance(item, dict)]
    require(any(all(item.get(field) == full_suite_descriptor.get(field)
                   for field in ("path", "bytes", "sha256"))
                for _, item in full_bindings),
            "quality-v3 receipt does not retain the old full suite")
    amendment_bindings = [(key, item) for key, item in quality_value.items()
                          if "amend" in key and isinstance(item, dict)]
    if amendment_bindings:
        same_descriptor(amendment_bindings[0][1],
                        descriptor(PACKET / "active-stage-v3.json",
                                  "v3 amendment"),
                        "quality-v3 amendment")
    quality_rows = [read_command("quality-v3-" + gate, stage,
                                 quality_v3_argv(gate, test_scope), 0)
                    for gate in gates]

    qvalue = read_json(qualification_path)
    require(qvalue.get("status") == "commands_pass"
            and isinstance(qvalue.get("rows"), list)
            and len(qvalue["rows"]) == 7,
            "qualification-v3 receipt changed")
    qlabels = [f"qualification-v3-{index:02}" for index in range(6)]
    qlabels.append("qualification-v3-opc-pair")
    qrows = []
    retained_reports = 0
    retained_samples = 0
    for index, label in enumerate(qlabels):
        case = CASES[index] if index < 6 else ",".join(CASES[2:4])
        qrows.append(validate_workload_row(
            qvalue["rows"][index], label, stage, build_info["binary"],
            case, "warm,cold-verified", 1, 0))
        retained_reports += 1
        retained_samples += count_qualification_report(
            PACKET / f"{label}.json", label, 4 if index == 6 else 2)
    require(retained_reports == 7 and retained_samples == 16,
            "qualification-v3 retained counts changed")
    validate_serial([build_command] + quality_rows + qrows)
    return {"present": True, "stage": stage, "admitted": True,
            "freeze": freeze, "build": build_info,
            "build_command": build_command,
            "quality": {"path": str(quality_path.relative_to(PACKET)),
                        "bytes": quality_path.stat().st_size,
                        "sha256": sha256(quality_path),
                        "test_scope": test_scope,
                        "commands": quality_rows},
            "qualification": {
                "path": str(qualification_path.relative_to(PACKET)),
                "bytes": qualification_path.stat().st_size,
                "sha256": sha256(qualification_path),
                "commands": qrows, "status": "commands_pass",
                "retained_reports": retained_reports,
                "retained_samples": retained_samples},
            "commands": [build_command] + quality_rows + qrows}


def expected_plan_rows() -> list[dict[str, Any]]:
    rows = []
    for block in range(6):
        order = CASES if block % 2 == 0 else tuple(reversed(CASES))
        states = ("warm", "cold-verified") if block % 2 == 0 else (
            "cold-verified", "warm")
        for state in states:
            rows.extend({"block": block, "case": case, "cache_state": state,
                          "samples": 30, "warmup": 3}
                         for case in order)
    return rows


def nearest_rank(values: list[int | float], quantile: float) -> int | float:
    require(values, "formal analysis has an empty sample vector")
    ordered = sorted(values)
    return ordered[max(0, math.ceil(len(ordered) * quantile) - 1)]


def sample_summary(values: list[int | float]) -> dict[str, Any]:
    require(values, "formal analysis has an empty metric vector")
    return {"min": min(values), "p50": nearest_rank(values, .5),
            "p95": nearest_rank(values, .95),
            "p99": nearest_rank(values, .99), "max": max(values),
            "mean": statistics.mean(values)}


def spread_summary(values: list[int | float]) -> dict[str, Any]:
    require(values, "formal analysis has an empty block vector")
    minimum = min(values)
    return {"min": minimum, "median": statistics.median(values),
            "max": max(values),
            "max_over_min": (max(values) / minimum if minimum > 0 else None)}


def recompute_formal_analysis(capture: dict[str, Any]) -> dict[str, Any]:
    """Recompute analyze.py's descriptive result from captured reports."""
    groups: dict[tuple[str, str], list[dict[str, Any]]] = {}
    for item in capture["rows"]:
        plan = item["plan"]
        result_binding = item["result"]["report"]
        report_path = packet_path(result_binding["path"], "formal report")
        report = read_json(report_path)
        reports = report.get("results")
        evidence = report.get("filesystem_evidence")
        require(isinstance(reports, list) and len(reports) == 1
                and isinstance(evidence, list) and len(evidence) == 1,
                "formal report result/evidence count changed")
        measured = reports[0]
        filesystem = evidence[0]
        samples = filesystem.get("samples")
        elapsed = measured.get("elapsed_ns")
        require(isinstance(samples, list) and len(samples) == 30
                and isinstance(elapsed, dict)
                and elapsed.get("samples") == [sample.get("elapsed_ns")
                                                for sample in samples],
                "formal sample vectors changed")
        times = elapsed["samples"]
        require(all(type(value) is int and value >= 0 for value in times),
                "formal elapsed sample is invalid")
        require(measured.get("case") == plan["case"]
                and measured.get("cache_state") == plan["cache_state"],
                "formal report case/state binding changed")
        first_metrics = samples[0].get("process_metrics")
        require(isinstance(first_metrics, dict),
                "formal process metrics are missing")
        metric_names = [name for name in first_metrics
                        if name != "clock_ticks_per_second"]
        metrics: dict[str, dict[str, Any]] = {}
        for name in metric_names:
            values = []
            for sample in samples:
                current = sample.get("process_metrics")
                require(isinstance(current, dict) and name in current,
                        f"formal process metric changed: {name}")
                values.append(current[name])
            require(all(type(value) in (int, float)
                        and not isinstance(value, bool) for value in values),
                    f"formal process metric is invalid: {name}")
            metrics[name] = sample_summary(values)
        key = (plan["case"], plan["cache_state"])
        groups.setdefault(key, []).append({
            "block": plan["block"], "latency_ns": sample_summary(times),
            "process_metrics": metrics, "report_sha256": result_binding["sha256"]})

    expected_keys = {(case, state) for case in CASES
                     for state in ("warm", "cold-verified")}
    require(set(groups) == expected_keys and all(len(rows) == 6
                                                  for rows in groups.values()),
            "formal analysis group counts changed")
    distributions = []
    lookup: dict[tuple[str, str], dict[str, Any]] = {}
    flags = []
    for case, state in sorted(groups):
        rows = sorted(groups[(case, state)], key=lambda row: row["block"])
        require([row["block"] for row in rows] == list(range(6)),
                f"formal block order changed: {case}/{state}")
        latency = {name: spread_summary([row["latency_ns"][name]
                                          for row in rows])
                   for name in rows[0]["latency_ns"]}
        process = {name: {quantile: spread_summary(
                      [row["process_metrics"][name][quantile] for row in rows])
                          for quantile in rows[0]["process_metrics"][name]}
                   for name in rows[0]["process_metrics"]}
        for metric in ("p50", "p95", "p99"):
            if latency[metric]["max_over_min"] is not None \
                    and latency[metric]["max_over_min"] > 1.2:
                flags.append({"case": case, "state": state, "metric": metric,
                              "reason": "block spread exceeds 20%",
                              "spread": latency[metric]["max_over_min"]})
        row = {"case": case, "cache_state": state, "blocks": rows,
               "latency_ns": latency, "process_metrics": process}
        distributions.append(row)
        lookup[(case, state)] = row

    pairs = []
    pairings = [(CASES[0], CASES[1]), (CASES[2], CASES[3]),
                (CASES[4], CASES[5])]
    for eager, source in pairings:
        for state in ("warm", "cold-verified"):
            eager_rows = lookup[(eager, state)]["blocks"]
            source_rows = lookup[(source, state)]["blocks"]
            ratios = [eager_row["latency_ns"]["p50"]
                      / source_row["latency_ns"]["p50"]
                      for eager_row, source_row in zip(eager_rows, source_rows)]
            require(all(value > 0 for value in ratios),
                    "formal paired ratio is invalid")
            rng = random.Random(834083)
            bootstrap = sorted(statistics.median(rng.choices(ratios, k=6))
                               for _ in range(10000))
            pairs.append({
                "eager": eager, "source": source, "cache_state": state,
                "per_block_eager_over_source_p50": ratios,
                "median": statistics.median(ratios),
                "bootstrap_95": [bootstrap[249], bootstrap[9749]],
                "scope": "descriptive route ratio; not an optimization effect"})
    return {
        "schema": "litchi.0834.descriptive-analysis.v1",
        "reports": 72, "samples": 2160, "bootstrap_seed": 834083,
        "bootstrap_resamples": 10000, "distributions": distributions,
        "route_ratios": pairs, "spread_flags": flags,
        "limitations": [
            "Synthetic fixed corpora and one host only.",
            "OPC eager open drops its package inside the timer; source open retains it through post-timer diagnostics.",
            "PPTX logical source counters come from an untimed replay.",
            "Verified cold means observed page-cache residency and process read_bytes, not physical-device I/O.",
            "Route/cache baseline only; no production optimization or before/after claim."]}


def validate_formal_analysis(capture: dict[str, Any]) -> dict[str, Any]:
    path = PACKET / "analysis.json"
    require(path.is_file() and not path.is_symlink(),
            "capture exists without independent analysis.json")
    expected = recompute_formal_analysis(capture)
    require(read_json(path) == expected,
            "formal analysis does not independently recompute")
    return {"analysis": source_descriptor(path, "formal analysis"),
            "analyzer": source_descriptor(PACKET / "analyze.py",
                                           "formal analyzer")}


def validate_optional_formal(repaired: dict[str, Any]) -> dict[str, Any]:
    plan_path = PACKET / "measurement-plan.json"
    if not plan_path.is_file():
        return {}
    plan = read_json(plan_path)
    require(plan.get("schema") == PLAN_SCHEMA and plan.get("base") == BASE
            and plan.get("cpu") == 12
            and plan.get("expected_reports") == 72
            and plan.get("expected_samples") == 2160
            and plan.get("rows") == expected_plan_rows(),
            "measurement plan changed")
    statistics_plan = plan.get("statistics")
    require(isinstance(statistics_plan, dict)
            and statistics_plan.get("process_quantiles") == "nearest rank"
            and statistics_plan.get("summary")
            == "midpoint median of six process p50 values"
            and statistics_plan.get("paired_comparisons")
            == "source/eager per operation and cache state; block-paired ratios"
            and statistics_plan.get("bootstrap_resamples") == 10000
            and statistics_plan.get("seed") == 833833
            and statistics_plan.get("sorted_endpoints") == [250, 9749]
            and statistics_plan.get("spread_flag")
            == "max/min process p50 or RSS exceeds 1.05"
            and statistics_plan.get("tail_flag")
            == "per-process p99/p50 exceeds 1.05",
            "measurement statistics plan changed")
    result = {"plan": source_descriptor(plan_path, "measurement plan")}
    admission_paths = [PACKET / "capture-admission.json",
                       PACKET / "admission.json"]
    existing_admissions = [path for path in admission_paths if path.is_file()]
    for admission_path in existing_admissions:
        admission = read_json(admission_path)
        require(admission.get("status") == "pass"
                and isinstance(admission.get("inputs"), dict)
                and admission["inputs"],
                f"{admission_path.name} receipt changed")
        for raw, item in admission["inputs"].items():
            require(isinstance(raw, str) and raw
                    and not Path(raw).is_absolute(),
                    f"{admission_path.name} input path is invalid")
            input_path = packet_path(raw, "admission input")
            same_descriptor(item, descriptor(input_path, "admission input"),
                            f"{admission_path.name} input {raw}")
        result[admission_path.stem] = source_descriptor(
            admission_path, admission_path.stem)
    capture_path = PACKET / "capture.json"
    if not capture_path.is_file():
        require(not (PACKET / "capture-failed.json").exists(),
                "formal capture has a retained failed attempt")
        return result
    require(repaired.get("present") is True,
            "capture exists without a validated repaired stage")
    require(repaired.get("admitted") is True,
            "capture exists without an admitted repaired qualification")
    stage = repaired.get("stage", "repaired")
    require(stage in ("repaired", "repaired-v2"),
            "formal capture stage changed")
    require(existing_admissions,
            "capture exists without an admission receipt")
    require(not (PACKET / "capture-failed.json").exists(),
            "capture has both pass and failed receipts")
    capture = read_json(capture_path)
    require(capture.get("status") == "commands_pass"
            and capture.get("report_count") == 72
            and capture.get("sample_count") == 2160
            and isinstance(capture.get("rows"), list)
            and len(capture["rows"]) == 72,
            "capture receipt changed")
    capture_start = PACKET / "capture-started.json"
    start = read_json(capture_start)
    same_descriptor(start.get("plan"),
                    descriptor(plan_path, "capture plan"),
                    "capture plan")
    admission_path = PACKET / "admission.json"
    if not admission_path.is_file():
        admission_path = PACKET / "capture-admission.json"
    same_descriptor(start.get("admission"),
                    descriptor(admission_path, "capture admission"),
                    "capture admission")
    capture_driver_binding = start.get("driver")
    require(isinstance(capture_driver_binding, dict),
            "capture driver binding is malformed")
    capture_driver_path = packet_path(capture_driver_binding.get("path"),
                                      "capture driver")
    same_descriptor(start.get("driver"),
                    descriptor(capture_driver_path, "capture driver"),
                    "capture driver")
    same_descriptor(start.get("freeze"),
                    descriptor(expected_freeze(stage), "capture freeze"),
                    "capture freeze")
    capture_commands = []
    for index, row in enumerate(capture["rows"]):
        expected = expected_plan_rows()[index]
        require(row.get("plan") == expected,
                f"capture plan row changed: {index}")
        label = f"native-{index:03}"
        report = PACKET / f"{label}.json"
        capture_commands.append(validate_workload_row(
            row.get("result"), label, stage, repaired["build"]["binary"],
            expected["case"], expected["cache_state"], 30, 3))
    validate_serial(capture_commands)
    result["capture"] = {"path": str(capture_path.relative_to(PACKET)),
                         "bytes": capture_path.stat().st_size,
                         "sha256": sha256(capture_path),
                         "commands": capture_commands}
    result["analysis"] = validate_formal_analysis(capture)
    return result


def encoded(value: dict[str, Any]) -> bytes:
    text = json.dumps(value, indent=2, sort_keys=True) + "\n"
    require(len(text.encode("utf-8")) < 5 * 1024 * 1024,
            "audit output is unexpectedly large")
    return text.encode("utf-8")


def write_or_check(value: dict[str, Any], check: bool, preview: bool) -> None:
    path = PACKET / "audit.json"
    data = encoded(value)
    if preview:
        return
    if check or path.exists():
        require(path.is_file() and not path.is_symlink(),
                "audit.json is missing")
        require(path.read_bytes() == data,
                "audit.json does not replay deterministically")
        return
    try:
        with path.open("xb") as stream:
            stream.write(data)
    except FileExistsError:
        fail("refusing to overwrite retained audit.json")


def build_audit() -> dict[str, Any]:
    origin = load_origin()
    custody = validate_source_archives(origin)
    host = validate_host(origin)
    build_info, build_command = validate_build_diagnostic()
    diagnostic = validate_diagnostic(build_info, build_command)
    all_commands = [build_command,
                    {**read_command("diagnostic-pptx-warm", "diagnostic",
                                    diagnostic_argv(build_info["binary"]["path"],
                                                    "warm",
                                                    PACKET / "diagnostic-pptx-warm.json"), 0)},
                    {**read_command("diagnostic-pptx-cold-verified", "diagnostic",
                                    diagnostic_argv(build_info["binary"]["path"],
                                                    "cold-verified",
                                                    PACKET / "diagnostic-pptx-cold-verified.json"), 1)}]
    active = validate_active_stage()
    if active.get("present"):
        all_commands.extend(active["commands"])
    repaired_v2 = validate_optional_repaired(active)
    follow = validate_follow_stage()
    repaired_v3 = validate_optional_repaired_v3(follow, repaired_v2)
    repaired = repaired_v3 if repaired_v3.get("present") else repaired_v2
    reader_tests = validate_optional_pptx_reader_tests()
    formal = validate_optional_formal(repaired)
    if repaired_v2.get("present"):
        all_commands.extend(repaired_v2["commands"])
    if repaired_v3.get("present"):
        all_commands.extend(repaired_v3["commands"])
    if formal.get("capture"):
        all_commands.extend(formal["capture"]["commands"])
    validate_command_inventory(all_commands)
    validate_serial(all_commands)
    return {
        "schema": AUDIT_SCHEMA,
        "status": "pass",
        "base": BASE,
        "custody": custody,
        "host": host,
        "build": build_info,
        "diagnostic": diagnostic,
        "active_stage": active,
        "commands": {
            "count": len(all_commands),
            "serial": True,
            "rows": all_commands,
        },
        "repaired": repaired,
        "repaired_v2": repaired_v2,
        "repaired_v3": repaired_v3,
        "pptx_reader_tests": reader_tests,
        "formal": formal,
        "claims": {
            "performance_claim": "none",
            "diagnostic": "failed untimed verified-prime replay independently decoded and recomputed",
            "formal_capture": "optional until repaired quality, qualification, admission, and plan exist",
        },
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    modes = parser.add_mutually_exclusive_group()
    modes.add_argument("--check", action="store_true",
                       help="replay and compare retained audit.json")
    modes.add_argument("--preview", action="store_true",
                       help="run validations without creating audit.json")
    args = parser.parse_args(argv)
    try:
        write_or_check(build_audit(), args.check, args.preview)
    except (AuditError, AssertionError, OSError, UnicodeError, ValueError,
            KeyError, TypeError, IndexError) as error:
        print(f"0834 independent audit failed: {error}", file=sys.stderr)
        return 1
    print(json.dumps({"status": "pass", "diagnostic": "retained",
                      "formal": bool((PACKET / "capture.json").is_file())},
                     sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
