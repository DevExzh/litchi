#!/usr/bin/env python3
"""Prepare the external, prospective 0966 DOC SPRM census plan.

This helper only hashes already existing files and writes a bounded JSON plan.
It never builds a binary, starts a child, opens a fixture, or changes the
source export.  The binary descriptor is therefore required at plan time so
the eventual runner can fail closed on any executable drift.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import os
from pathlib import Path
import re
from typing import Any


ARTIFACT_ROOT = Path(
    os.environ.get("LITCHI_PERF_0966_ROOT", "/home/zhuhe/litchi-performance-0966")
).resolve()
SOURCE_DEFAULT = ARTIFACT_ROOT / "source-head-v1"
PLAN_SCHEMA = "litchi.performance.0966.doc_sprm_census_plan.v1"
REPORT_SCHEMA = "litchi.performance.0966.doc_sprm_census.v1"
CASES = ("docnohf", "docfloat")
CASE_FILES = {
    "docnohf": "test-data/ole/doc/NoHeadFoot.doc",
    "docfloat": "test-data/ole/doc/FloatingPictures.doc",
}
EDIT_COUNTS = (1, 2, 8, 32)
RETENTIONS = ("zero", "default_8mib", "release_8mib")
RETENTION_ARG = {
    "zero": "zero",
    "default_8mib": "default",
    "release_8mib": "release",
}
ENVIRONMENT = {
    "PATH": "/usr/bin:/bin",
    "LC_ALL": "C",
    "LANG": "C",
    "TZ": "UTC",
    "RUST_BACKTRACE": "0",
    "PAGER": "cat",
}
MAX_JSON_BYTES = 1 << 20
REVISION_RE = re.compile(r"^[0-9a-f]{40}$")


class PlanError(RuntimeError):
    pass


def require(condition: bool, message: str) -> None:
    if not condition:
        raise PlanError(message)


def utc_now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat().replace("+00:00", "Z")


def regular_file(path: Path, label: str) -> Path:
    path = path.resolve()
    require(path.is_file() and not path.is_symlink(), f"{label}: not a regular file: {path}")
    return path


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with regular_file(path, "hash").open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def descriptor(path: Path, label: str) -> dict[str, Any]:
    path = regular_file(path, label)
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256_file(path)}


def source_tree_descriptor(root: Path) -> dict[str, Any]:
    root = root.resolve()
    require(root.is_dir() and not root.is_symlink(), f"source root is not a directory: {root}")
    entries: list[tuple[str, int, str]] = []
    for current, directories, files in os.walk(root, topdown=True, followlinks=False):
        current_path = Path(current)
        directories.sort()
        files.sort()
        for name in files:
            path = current_path / name
            regular_file(path, "source tree")
            relative = path.relative_to(root).as_posix()
            size = path.stat().st_size
            entries.append((relative, size, sha256_file(path)))
        for name in directories:
            path = current_path / name
            require(not path.is_symlink(), f"source tree contains symlink directory: {path}")
    digest = hashlib.sha256()
    for relative, size, file_hash in entries:
        digest.update(relative.encode("utf-8"))
        digest.update(b"\0")
        digest.update(size.to_bytes(8, "big"))
        digest.update(file_hash.encode("ascii"))
        digest.update(b"\n")
    return {
        "root": str(root),
        "file_count": len(entries),
        "tree_sha256": digest.hexdigest(),
    }


def json_descriptor(path: Path, label: str) -> dict[str, Any]:
    value = descriptor(path, label)
    require(value["bytes"] < MAX_JSON_BYTES, f"{label}: JSON/control file is too large")
    return value


def review_status(path: Path) -> str:
    """Require the independently authored review JSON to report a pass."""
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise PlanError(f"doc_census_review must be passing JSON: {error}") from error
    require(isinstance(value, dict), "doc_census_review is not a JSON object")
    status = value.get("status")
    require(status in {"pass", "passed"}, "doc_census_review status is not pass/passed")
    return status


def harness_files(source_root: Path) -> list[dict[str, Any]]:
    relative = [
        "Cargo.toml",
        "Cargo.lock",
        "README.md",
        "src/alloc_metrics.rs",
        "src/bin/alloc.rs",
        "src/lib.rs",
        "src/main.rs",
        "src/oracle_0730.rs",
    ]
    return [descriptor(source_root / "tools/perf-cfb-profile" / item, f"harness {item}") for item in relative]


def cell_rows() -> list[dict[str, Any]]:
    cells = [
        (case, edits, retention)
        for case in CASES
        for edits in EDIT_COUNTS
        for retention in RETENTIONS
    ]
    rows: list[dict[str, Any]] = []
    position = 0
    for repeat, direction, ordered in ((1, "forward", cells), (2, "reverse", list(reversed(cells)))):
        for index, (case, edits, retention) in enumerate(ordered, start=1):
            position += 1
            rows.append(
                {
                    "id": f"sprm-r{repeat:02d}-{case}-e{edits:02d}-{retention}",
                    "position": position,
                    "repeat": repeat,
                    "direction": direction,
                    "repeat_position": index,
                    "case": case,
                    "edits": edits,
                    "retention": retention,
                    "retention_arg": RETENTION_ARG[retention],
                    "warmups": 0,
                    "samples": 1,
                    "retry": False,
                }
            )
    return rows


def write_exclusive(path: Path, value: object) -> None:
    path = path.resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    data = (json.dumps(value, indent=2, sort_keys=True) + "\n").encode("utf-8")
    require(len(data) < MAX_JSON_BYTES, f"plan exceeds {MAX_JSON_BYTES} bytes")
    try:
        with path.open("xb") as stream:
            stream.write(data)
            stream.flush()
            os.fsync(stream.fileno())
    except FileExistsError as error:
        raise PlanError(f"refusing to overwrite retained plan: {path}") from error


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-root", type=Path, default=SOURCE_DEFAULT)
    parser.add_argument("--binary", type=Path, required=True)
    parser.add_argument("--review", type=Path, required=True,
                        help="independent doc_census_review document")
    parser.add_argument("--checker", type=Path,
                        default=ARTIFACT_ROOT / "check_sprm_census_0966.py")
    parser.add_argument("--runner", type=Path,
                        default=ARTIFACT_ROOT / "tools/run_doc_sprm_census_0966_v1.py")
    parser.add_argument("--reducer", type=Path,
                        default=ARTIFACT_ROOT / "tools/reduce_doc_sprm_census_0966_v1.py")
    parser.add_argument("--source-revision", default="a499d34f5119a034bc92abde1e1dfcf529159b8d")
    parser.add_argument("--output", type=Path,
                        default=ARTIFACT_ROOT / "protocol/plan-doc-sprm-census-0966-v1.json")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    source_root = args.source_root.resolve()
    artifact_root = ARTIFACT_ROOT
    tree = source_tree_descriptor(source_root)
    parser_source = descriptor(source_root / "crates/litchi-doc/src/sprm.rs", "SPRM parser")
    fixture_descriptors = {
        case: descriptor(source_root / relative, f"fixture {case}")
        for case, relative in CASE_FILES.items()
    }
    binary = descriptor(args.binary, "census executable")
    binary_path = Path(binary["path"])
    require(os.access(binary_path, os.X_OK), f"census executable is not executable: {binary_path}")
    review = json_descriptor(args.review, "doc_census_review")
    review_result = review_status(args.review)
    checker = json_descriptor(args.checker, "checker")
    tools = {
        "runner": json_descriptor(args.runner, "runner"),
        "reducer": json_descriptor(args.reducer, "reducer"),
    }
    require(REVISION_RE.fullmatch(args.source_revision) is not None,
            "source revision must be exactly 40 lowercase hexadecimal characters")
    rows = cell_rows()
    require(len(rows) == 48, "schedule must contain exactly 48 rows")
    plan = {
        "schema": PLAN_SCHEMA,
        "schema_version": 1,
        "change": "0966",
        "created_utc": utc_now(),
        "status": "sealed_ready_for_go",
        "status_note": "This sealed plan only binds custody and schedule; root GO is required before capture.",
        "source_revision": args.source_revision,
        "claim_scope": {
            "purpose": "bounded DOC SPRM parser census",
            "diagnostic_only": True,
            "timing_claim": False,
            "allocator_total_claim": False,
            "elapsed_time_claim": False,
            "removable_work_claim": False,
        },
        "source": {
            **tree,
            "parser": parser_source,
            "harness_files": harness_files(source_root),
        },
        "fixtures": fixture_descriptors,
        "executable": binary,
        "review": {
            "role": "doc_census_review",
            "status": "bound_external_independent_review_passed",
            "document_status": review_result,
            "document": review,
        },
        "checker": checker,
        "tools": tools,
        "run_policy": {
            "serial": True,
            "retry": False,
            "warmups": 0,
            "samples": 1,
            "deadline_seconds": 300,
            "max_stdout_bytes": 1 << 20,
            "max_stderr_bytes": 1 << 20,
            "terminal_before_parse": True,
            "raw_stdout_stderr_retained": True,
            "cpu_affinity": None,
        },
        "environment": ENVIRONMENT,
        "capture_root": str(artifact_root / "capture/doc-sprm-census-v1"),
        "reduction_output": str(artifact_root / "analysis/doc-sprm-census-reduction-v1.json"),
        "report_schema": REPORT_SCHEMA,
        "argv_template": [
            binary["path"],
            "--case", "<case>",
            "--input", "<fixture>",
            "--edits", "<edits>",
            "--retention", "<retention_arg>",
            "--warmup", "0",
            "--samples", "1",
        ],
        "schedule": {
            "cell_dimensions": {
                "cases": list(CASES),
                "edits": list(EDIT_COUNTS),
                "retentions": list(RETENTIONS),
            },
            "repeat_directions": ["forward", "reverse"],
            "rows": rows,
            "children": len(rows),
            "cells": 24,
        },
    }
    write_exclusive(args.output, plan)
    print(json.dumps({"plan": str(args.output.resolve()), "rows": len(rows), "status": plan["status"]}, sort_keys=True))


if __name__ == "__main__":
    try:
        main()
    except PlanError as error:
        raise SystemExit(f"FAIL: {error}")
