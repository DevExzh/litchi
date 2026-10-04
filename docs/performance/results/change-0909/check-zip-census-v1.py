#!/usr/bin/env python3
"""Offline checks for the 0909 raw ZIP census.

This checker creates small deterministic ZIPs in a temporary directory and
invokes the sibling census as a subprocess.  It exercises stored and Deflate
members, an explicitly allowed selected-slide edit, and rejection of an
untouched payload, metadata, order, or archive-comment change.  It deliberately
does not claim ZIP64 or data-descriptor coverage.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
from typing import Any, Sequence
import zipfile


HERE = Path(__file__).resolve().parent
CENSUS = HERE / "zip-census.py"
ALLOW = "ppt/slides/slide1.xml"
SELECTED = b"<p:sld><a:t>selected replacement</a:t></p:sld>"
UNSELECTED = b"<p:sld><a:t>untouched source</a:t></p:sld>"


class CheckError(RuntimeError):
    pass


def require(condition: bool, message: str) -> None:
    if not condition:
        raise CheckError(message)


def write_archive(path: Path, members: Sequence[tuple[str, bytes]], compression: int, comment: bytes = b"fixture-comment") -> None:
    with zipfile.ZipFile(path, "w", compression=compression) as archive:
        archive.comment = comment
        for name, data in members:
            info = zipfile.ZipInfo(name, (2020, 1, 2, 3, 4, 5))
            info.compress_type = compression
            info.create_system = 3
            info.external_attr = 0o600 << 16
            archive.writestr(info, data)


def invoke(root: Path, label: str, source: Path, output: Path, *, allowed: Sequence[str] = ()) -> tuple[subprocess.CompletedProcess[str], dict[str, Any]]:
    report = root / f"{label}.json"
    command = [
        sys.executable,
        str(CENSUS),
        "--source",
        str(source),
        "--output",
        str(output),
        "--report",
        str(report),
    ]
    for name in allowed:
        command.extend(("--allow-changed", name))
    result = subprocess.run(command, capture_output=True, text=True, check=False)
    require(report.is_file(), f"{label}: census did not create its report")
    try:
        receipt = json.loads(report.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise CheckError(f"{label}: invalid census report: {error}") from error
    return result, receipt


def check_positive(root: Path, compression: int, label: str) -> dict[str, Any]:
    source = root / f"{label}-source.pptx"
    output = root / f"{label}-output.pptx"
    members = [(ALLOW, b"<p:sld><a:t>selected source</a:t></p:sld>"), ("ppt/slides/slide2.xml", UNSELECTED), ("ppt/media/image.bin", bytes(range(64)))]
    write_archive(source, members, compression)
    write_archive(output, members, compression)
    result, receipt = invoke(root, label, source, output, allowed=(ALLOW,))
    require(result.returncode == 0, f"{label}: expected success, got {result.returncode}: {result.stderr}")
    require(receipt.get("status") == "pass", f"{label}: report is not a pass")
    require(receipt["archive"]["source"]["data_descriptors"] is False, f"{label}: unexpected data descriptor")
    require(receipt["archive"]["source"]["zip64"] is False, f"{label}: unexpected ZIP64")
    require(receipt["changed_members"]["names"] == [], f"{label}: identical archive changed members")
    return {"label": label, "kind": "positive_identical", "compression": compression, "returncode": result.returncode}


def check_allowed_edit(root: Path, compression: int, label: str) -> dict[str, Any]:
    source = root / f"{label}-source.pptx"
    output = root / f"{label}-output.pptx"
    write_archive(source, [(ALLOW, b"<p:sld><a:t>selected source</a:t></p:sld>"), ("ppt/slides/slide2.xml", UNSELECTED)], compression)
    write_archive(output, [(ALLOW, SELECTED), ("ppt/slides/slide2.xml", UNSELECTED)], compression)
    result, receipt = invoke(root, label, source, output, allowed=(ALLOW,))
    require(result.returncode == 0, f"{label}: allowed edit failed: {result.stderr}")
    require(receipt.get("status") == "pass", f"{label}: allowed edit report is not a pass")
    require(receipt["changed_members"]["names"] == [ALLOW], f"{label}: changed list is not selected-only")
    require(receipt["members"][1]["changed_components"] == [], f"{label}: untouched member changed")
    return {"label": label, "kind": "positive_allowed_selected_edit", "compression": compression, "returncode": result.returncode}


def check_rejection(root: Path, label: str, source_members: Sequence[tuple[str, bytes]], output_members: Sequence[tuple[str, bytes]], *, source_compression: int = zipfile.ZIP_DEFLATED, output_compression: int | None = None, source_comment: bytes = b"fixture-comment", output_comment: bytes | None = None, expected_reason: str, expected_component: str | None = None) -> dict[str, Any]:
    source = root / f"{label}-source.pptx"
    output = root / f"{label}-output.pptx"
    if output_compression is None:
        output_compression = source_compression
    if output_comment is None:
        output_comment = source_comment
    write_archive(source, source_members, source_compression, source_comment)
    write_archive(output, output_members, output_compression, output_comment)
    result, receipt = invoke(root, label, source, output, allowed=(ALLOW,))
    require(result.returncode == 1, f"{label}: expected comparison rejection, got {result.returncode}: {result.stderr}")
    require(receipt.get("status") == "fail", f"{label}: expected fail report")
    reasons = {violation.get("reason") for violation in receipt["violations"]}
    require(expected_reason in reasons, f"{label}: missing violation reason {expected_reason}: {reasons}")
    if expected_component is not None:
        changed = receipt["changed_members"]["by_component"].get(expected_component, [])
        require("ppt/slides/slide2.xml" in changed, f"{label}: missing component {expected_component}: {changed}")
    return {"label": label, "kind": "negative_rejection", "reason": expected_reason, "returncode": result.returncode}


def run_checks() -> dict[str, Any]:
    root = Path(tempfile.mkdtemp(prefix="litchi-0909-zip-census-check-"))
    try:
        results = [
            check_positive(root, zipfile.ZIP_STORED, "stored-identical"),
            check_positive(root, zipfile.ZIP_DEFLATED, "deflated-identical"),
            check_allowed_edit(root, zipfile.ZIP_STORED, "stored-selected-edit"),
            check_allowed_edit(root, zipfile.ZIP_DEFLATED, "deflated-selected-edit"),
            check_rejection(
                root,
                "untouched-payload",
                [(ALLOW, b"selected"), ("ppt/slides/slide2.xml", UNSELECTED)],
                [(ALLOW, b"selected replacement"), ("ppt/slides/slide2.xml", b"changed untouched payload")],
                expected_reason="changed_untouched_member",
                expected_component="decoded_payload",
            ),
            check_rejection(
                root,
                "untouched-metadata",
                [(ALLOW, b"selected"), ("ppt/slides/slide2.xml", UNSELECTED)],
                [(ALLOW, b"selected replacement"), ("ppt/slides/slide2.xml", UNSELECTED)],
                output_compression=zipfile.ZIP_STORED,
                expected_reason="changed_untouched_member",
                expected_component="central_metadata",
            ),
            check_rejection(
                root,
                "untouched-order",
                [(ALLOW, b"selected"), ("ppt/slides/slide2.xml", UNSELECTED)],
                [("ppt/slides/slide2.xml", UNSELECTED), (ALLOW, b"selected")],
                expected_reason="ordered_filenames_changed",
            ),
            check_rejection(
                root,
                "archive-comment",
                [(ALLOW, b"selected"), ("ppt/slides/slide2.xml", UNSELECTED)],
                [(ALLOW, b"selected replacement"), ("ppt/slides/slide2.xml", UNSELECTED)],
                output_comment=b"changed-comment",
                expected_reason="archive_comment_changed",
            ),
        ]
        return {
            "schema": "litchi.performance.zip-census-check.v1",
            "status": "pass",
            "cases": results,
            "coverage": {
                "stored": True,
                "deflated": True,
                "zip64": False,
                "data_descriptors": False,
                "claim": "This checker does not claim ZIP64 or data-descriptor coverage.",
            },
        }
    finally:
        shutil.rmtree(root)


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.parse_args(argv)
    try:
        receipt = run_checks()
    except (CheckError, OSError, subprocess.SubprocessError) as error:
        print(f"error: {error}", file=sys.stderr)
        return 1
    print(json.dumps(receipt, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
