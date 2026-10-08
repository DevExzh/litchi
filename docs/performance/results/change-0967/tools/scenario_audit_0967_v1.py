#!/usr/bin/env python3
"""Audit the 0967 DOC scenario matrix before any timing capture.

This is a source/receipt audit. It never builds or launches a workload. The
control executable reports are supplied by the separate Rust probe; this tool
only checks their frozen semantics and binds the real fixture identities.
"""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
from typing import Any

SCHEMA = "litchi.performance.0967.scenario-audit.v1"
REAL = {
    "docnohf": {
        "relative_path": "test-data/ole/doc/NoHeadFoot.doc",
        "bytes": 26112,
        "sha256": "45e5df073f34314da6f39d2dad119fb2ef23470878fd2df67f632864cd92ea48",
    },
    "docfloat": {
        "relative_path": "test-data/ole/doc/FloatingPictures.doc",
        "bytes": 335360,
        "sha256": "41cc0cd8f2d7390266f844f5bddbcd29d08ac5338abc33832d9c9a6ed5d0f85d",
    },
}
QUERY = {
    "id": "real-bookmark-delete-redline-query",
    "relative_path": "test-data/libreoffice-core/sw/qa/core/doc/data/bookmark-delete-redline.doc",
    "bytes": 66048,
    "sha256": "f8c19b5712e12244af5e243ab6b996af6bc32193dad208948a155c4a4d5bb3dd",
    "revision_count": 20,
    "paragraph_zero_unmarked": False,
    "lifecycle": "query_reopen_refusal",
}
REQUIRED_CONTROLS = {
    "generated-positive": "lifecycle_revision_positive",
    "generated-dense-1024": "lifecycle_dense_marker",
    "generated-false": "lifecycle_false_marker",
    QUERY["id"]: QUERY["lifecycle"],
}


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def load(path: Path) -> dict[str, Any]:
    with path.open(encoding="utf-8") as stream:
        value = json.load(stream)
    if not isinstance(value, dict):
        raise ValueError(f"{path}: expected object")
    return value


def check_real_fixtures(root: Path) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for case, expected in REAL.items():
        path = root / expected["relative_path"]
        if not path.is_file():
            raise ValueError(f"missing real fixture {path}")
        actual_bytes = path.stat().st_size
        actual_sha = sha256(path)
        if (actual_bytes, actual_sha) != (expected["bytes"], expected["sha256"]):
            raise ValueError(
                f"{case}: fixture identity mismatch: "
                f"{actual_bytes}/{actual_sha} != {expected['bytes']}/{expected['sha256']}"
            )
        rows.append({"id": case, "path": str(path), **expected})
    return rows


def census_check(census_path: Path, root: Path) -> dict[str, Any]:
    census = load(census_path)
    if census.get("schema") != "litchi.performance.0967.doc_revision_positive_control.v1":
        raise ValueError(f"unexpected census schema in {census_path}")
    by_path = {row.get("path"): row for row in census.get("rows", [])}
    checks: dict[str, Any] = {}
    def census_row(relative_path: str) -> dict[str, Any] | None:
        return by_path.get(relative_path) or by_path.get(relative_path.removeprefix("test-data/"))
    for case, expected in REAL.items():
        row = census_row(expected["relative_path"])
        if row is None:
            raise ValueError(f"census lacks {expected['relative_path']}")
        checks[case] = {
            "revision_count": row.get("revision_count"),
            "paragraph_zero_unmarked": row.get("paragraph_zero_unmarked"),
            "error": row.get("error"),
        }
    row = census_row(QUERY["relative_path"])
    if row is None:
        raise ValueError(f"census lacks {QUERY['relative_path']}")
    if row.get("revision_count") != QUERY["revision_count"]:
        raise ValueError(f"query census revision count changed: {row}")
    if row.get("paragraph_zero_unmarked") is not QUERY["paragraph_zero_unmarked"]:
        raise ValueError(f"query census paragraph-zero state changed: {row}")
    checks[QUERY["id"]] = {
        "revision_count": row.get("revision_count"),
        "paragraph_zero_unmarked": row.get("paragraph_zero_unmarked"),
        "error": row.get("error"),
    }
    return {"path": str(census_path), "checks": checks}


def report_check(control_id: str, path: Path) -> dict[str, Any]:
    report = load(path)
    if report.get("schema") != "litchi.performance.0967.doc_revision_positive_control.v1":
        raise ValueError(f"{control_id}: bad report schema")
    if report.get("timing_claim") is not False:
        raise ValueError(f"{control_id}: control timing claim must be false")
    if report.get("error") is not None:
        raise ValueError(f"{control_id}: report error: {report['error']}")
    mode = report.get("mode")
    pre = report.get("pre") or {}
    post = report.get("post") or {}
    revisions_before = pre.get("revisions") or []
    revisions_after = post.get("revisions") or []
    if control_id == QUERY["id"]:
        if mode != "query":
            raise ValueError(f"{control_id}: expected query mode")
        query = report.get("query") or {}
        if len(revisions_before) != QUERY["revision_count"]:
            raise ValueError(f"{control_id}: expected 20 revisions")
        if query.get("replacement_succeeded") is not False:
            raise ValueError(f"{control_id}: paragraph-0 refusal was weakened")
        if not query.get("refusal") or query.get("source_unchanged_after_refusal") is not True:
            raise ValueError(f"{control_id}: refusal/reopen custody failed")
        return {
            "id": control_id,
            "path": str(path),
            "variant": report.get("variant"),
            "mode": mode,
            "revision_count": len(revisions_before),
            "lifecycle": QUERY["lifecycle"],
            "status": "pass",
        }
    if mode != "run":
        raise ValueError(f"{control_id}: expected lifecycle run mode")
    edit = report.get("edit") or {}
    if edit.get("replacement_succeeded") is not True or edit.get("reopen_succeeded") is not True:
        raise ValueError(f"{control_id}: replacement/reopen failed")
    if edit.get("changed") is not True:
        raise ValueError(f"{control_id}: lifecycle did not change output")
    if report.get("post", {}).get("paragraph_zero_text_sha256") == report.get("pre", {}).get(
        "paragraph_zero_text_sha256"
    ):
        raise ValueError(f"{control_id}: paragraph zero did not change")
    witness = report.get("branch_witness") or {}
    if control_id == "generated-false":
        if not str(report.get("fixture", "")).startswith("generated-false-marker"):
            raise ValueError(f"{control_id}: false-marker fixture identity missing")
        if witness.get("false_marker_present") is not True:
            raise ValueError(f"{control_id}: false-marker witness missing")
        if not revisions_before or revisions_after != revisions_before:
            raise ValueError(f"{control_id}: true revision control did not survive false marker route")
        lifecycle = "lifecycle_false_marker"
    else:
        if control_id == "generated-positive" and report.get("fixture") != "generated-positive":
            raise ValueError(f"{control_id}: generated fixture identity missing")
        if control_id == "generated-dense-1024" and report.get("fixture") != "generated-dense-1024":
            raise ValueError(f"{control_id}: dense fixture identity missing")
        if witness.get("positive_branch_required") is not True:
            raise ValueError(f"{control_id}: positive branch witness missing")
        if not revisions_before or revisions_after != revisions_before:
            raise ValueError(f"{control_id}: revision identities did not survive lifecycle")
        lifecycle = REQUIRED_CONTROLS[control_id]
    return {
        "id": control_id,
        "path": str(path),
        "variant": report.get("variant"),
        "mode": mode,
        "revision_count": len(revisions_before),
        "lifecycle": lifecycle,
        "status": "pass",
    }


def parse_receipts(values: list[str]) -> dict[str, list[dict[str, Any]]]:
    found: dict[str, list[dict[str, Any]]] = {}
    for value in values:
        try:
            control_id, raw_path = value.split("=", 1)
        except ValueError as exc:
            raise ValueError(f"receipt must be CONTROL_ID=PATH: {value}") from exc
        path = Path(raw_path)
        found.setdefault(control_id, []).append(report_check(control_id, path))
    return found


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-root", type=Path, required=True)
    parser.add_argument("--census", type=Path)
    parser.add_argument("--receipt", action="append", default=[], metavar="ID=PATH")
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    try:
        real = check_real_fixtures(args.source_root)
        census = census_check(args.census, args.source_root) if args.census else None
        receipts = parse_receipts(args.receipt)
        missing = [control_id for control_id in REQUIRED_CONTROLS if control_id not in receipts]
        incomplete = [
            control_id
            for control_id, reports in receipts.items()
            if control_id in REQUIRED_CONTROLS
            and not {report.get("variant") for report in reports}.issuperset({"baseline", "candidate"})
        ]
        status = "pass" if not missing and not incomplete else "pending_controls"
        result = {
            "schema": SCHEMA,
            "status": status,
            "source_root": str(args.source_root.resolve()),
            "real_fixtures": real,
            "census": census,
            "required_controls": REQUIRED_CONTROLS,
            "controls": receipts,
            "missing_controls": missing,
            "incomplete_controls": incomplete,
            "timing_claim": False,
            "build_started": False,
            "workload_started": False,
        }
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(json.dumps(result, indent=2) + "\n", encoding="utf-8")
        print(json.dumps({"status": status, "missing_controls": missing, "incomplete_controls": incomplete}))
        return 0 if status == "pass" else 2
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        print(f"scenario audit failed: {exc}")
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
