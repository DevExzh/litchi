#!/usr/bin/env python3
"""Verify the retained 0915 Callgrind artifacts and caller-scoped evidence.

The script reads only the small JSON packet in this directory and the frozen
or retained external artifacts named by that packet. It never rebuilds the
binary, writes to the external evidence directory, or overwrites a frozen
receipt. A fresh capture can be made with the command recorded in
``profile-reduced.json`` using new output names.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import re
import sys


class ProfileError(ValueError):
    """Raised when retained profile evidence does not match its packet."""


def check(condition: bool, message: str) -> None:
    if not condition:
        raise ProfileError(message)


def verify_file(metadata: dict, label: str) -> Path:
    path = Path(metadata["path"])
    check(path.is_file(), f"{label}: missing {path}")
    data = path.read_bytes()
    check(len(data) == metadata["bytes"], f"{label}: size mismatch for {path}")
    digest = hashlib.sha256(data).hexdigest()
    check(digest == metadata["sha256"], f"{label}: SHA-256 mismatch for {path}")
    return path


def function_records(path: Path) -> tuple[str, list[dict]]:
    """Read Callgrind cfn edges and the first event value after each edge."""

    text = path.read_text(errors="strict")
    lines = text.splitlines()
    function_names: dict[str, str] = {}
    for line in lines:
        match = re.match(r"fn=\((\d+)\)(?:\s+(.*))?$", line)
        if match is not None:
            name = match.group(2)
            if name:
                function_names[match.group(1)] = name
    records: list[dict] = []
    current_fn = ""
    pending_cfn: str | None = None
    for index, line in enumerate(lines):
        if line.startswith("fn="):
            match = re.match(r"fn=\((\d+)\)(?:\s+(.*))?$", line)
            current_fn = function_names.get(match.group(1), "") if match is not None else line
            pending_cfn = None
            continue
        if line.startswith("cfn="):
            match = re.match(r"cfn=\((\d+)\)(?:\s+(.*))?$", line)
            if match is not None:
                pending_cfn = match.group(2) or function_names.get(match.group(1), "")
            else:
                pending_cfn = line
            continue
        if pending_cfn is None or not line.startswith("calls="):
            continue
        match = re.match(r"calls=(\d+)", line)
        if match is None:
            continue
        cost = None
        for following in lines[index + 1 : index + 8]:
            event = re.match(r"^\s*\d+\s+(-?\d+)(?:\s|$)", following)
            if event is not None:
                cost = int(event.group(1))
                break
        records.append(
            {
                "caller": current_fn,
                "callee": pending_cfn,
                "calls": int(match.group(1)),
                "cost": cost,
            }
        )
        pending_cfn = None
    return text, records


def one_record(records: list[dict], predicate, label: str) -> dict:
    matches = [record for record in records if predicate(record)]
    check(len(matches) == 1, f"{label}: expected one record, found {len(matches)}")
    return matches[0]


def verify_owner(owner: dict) -> dict:
    path = Path(owner["dump"])
    text, records = function_records(path)
    part = int(re.search(r"^part:\s*(\d+)$", text, re.MULTILINE).group(1))
    expected_part = 16 if owner["owner"] == "query_snapshot" else 17
    check(part == expected_part, f"{owner['owner']}: part {part}, expected {expected_part}")

    if owner["owner"] == "query_snapshot":
        chain_fragments = (
            "main_section_layout_snapshot",
            "litchi_docx::source_backed::Package::section_layout_snapshot",
            "litchi_perf_baseline::run_docx_source_backed_existing_section_layout_edit_save",
        )
    else:
        chain_fragments = (
            "main_section_layout_snapshot",
            "litchi_docx::source_backed::Package::publish_section_layout_commit_to_stream",
            "litchi_perf_baseline::run_docx_source_backed_existing_section_layout_edit_save",
        )
    root = one_record(
        records,
        lambda record: all(fragment in record["callee"] for fragment in chain_fragments)
        and record["calls"] == 0,
        f"{owner['owner']}: selected caller chain",
    )
    check(root["cost"] == owner["selected_total_ir"], f"{owner['owner']}: selected Ir mismatch")

    snapshot = one_record(
        records,
        lambda record: "Snapshot::from_shared_xml" in record["caller"]
        and "main_section_layout_snapshot" in record["caller"]
        and record["callee"].startswith(
            "litchi_docx::source_backed::ensure_source_section_inventory_xml'"
        ),
        f"{owner['owner']}: snapshot/ensure caller",
    )
    check(snapshot["calls"] == owner["ensure_source_section_inventory_xml_calls"],
          f"{owner['owner']}: ensure call count mismatch")
    check(snapshot["cost"] == owner["ensure_ir"], f"{owner['owner']}: ensure Ir mismatch")

    parse = one_record(
        records,
        lambda record: "Snapshot::from_shared_xml" in record["caller"]
        and "main_section_layout_snapshot" in record["caller"]
        and "parse_inventory" in record["callee"],
        f"{owner['owner']}: snapshot/parse caller",
    )
    check(parse["calls"] == owner["parse_inventory_calls"],
          f"{owner['owner']}: parse call count mismatch")
    check(parse["cost"] == owner["parse_inventory_ir"], f"{owner['owner']}: parse Ir mismatch")

    sha = one_record(
        records,
        lambda record: "main_section_layout_snapshot" in record["callee"]
        and "sha2::sha256::compress256" in record["callee"]
        and record["cost"] == owner["sha256_compress256_ir"],
        f"{owner['owner']}: software SHA caller",
    )
    check(sha["cost"] == owner["sha256_compress256_ir"],
          f"{owner['owner']}: software SHA Ir mismatch")

    clone_records = {
        "process_event_calls": one_record(
            records,
            lambda record: record["callee"].startswith(
                "quick_xml::reader::ns_reader::NsReader<R>::process_event'"
            )
            and "ensure_source_section_inventory_xml" in record["callee"],
            f"{owner['owner']}: process-event clone context",
        )["calls"],
        "alloc_calls": one_record(
            records,
            lambda record: "__rustc::__rdl_alloc" in record["callee"]
            and "ensure_source_section_inventory_xml" in record["callee"],
            f"{owner['owner']}: clone allocation context",
        )["calls"],
        "dealloc_calls": one_record(
            records,
            lambda record: "__rustc::__rdl_dealloc" in record["callee"]
            and "ensure_source_section_inventory_xml" in record["callee"],
            f"{owner['owner']}: clone deallocation context",
        )["calls"],
        "memcpy_calls": one_record(
            records,
            lambda record: "__memcpy" in record["callee"]
            and "ensure_source_section_inventory_xml" in record["callee"],
            f"{owner['owner']}: clone copy context",
        )["calls"],
    }
    expected_clone = {
        "process_event_calls": 1838,
        "alloc_calls": 3676,
        "dealloc_calls": 3674,
        "memcpy_calls": 1838,
    }
    check(clone_records == expected_clone, f"{owner['owner']}: clone counts mismatch")
    return {
        "owner": owner["owner"],
        "part": part,
        "selected_total_ir": root["cost"],
        "ensure_ir": snapshot["cost"],
        "clone_counts": clone_records,
    }


def verify_repeat(repeat: dict, expected_part: int) -> dict:
    path = Path(repeat["dump"])
    text, records = function_records(path)
    part = int(re.search(r"^part:\s*(\d+)$", text, re.MULTILINE).group(1))
    check(part == expected_part, f"repeat {path}: part mismatch")
    owner = "section_layout_snapshot" if expected_part == 16 else "publish_section_layout_commit_to_stream"
    chain = f"main_section_layout_snapshot'litchi_docx::source_backed::Package::{owner}"
    root = one_record(
        records,
        lambda record: chain in record["callee"] and record["calls"] == 0,
        f"repeat {path}: caller chain",
    )
    check(root["cost"] == repeat["selected_total_ir"], f"repeat {path}: selected Ir mismatch")
    ensure = one_record(
        records,
        lambda record: "Snapshot::from_shared_xml" in record["caller"]
        and record["callee"].startswith(
            "litchi_docx::source_backed::ensure_source_section_inventory_xml'"
        ),
        f"repeat {path}: ensure caller",
    )
    check(ensure["cost"] == repeat["ensure_ir"], f"repeat {path}: ensure Ir mismatch")
    return {"path": str(path), "part": part, "selected_total_ir": root["cost"], "ensure_ir": ensure["cost"]}


def replay(packet: Path) -> dict:
    summary_path = packet / "profile-reduced.json"
    summary = json.loads(summary_path.read_text())
    checked = 0
    for label, metadata in summary["frozen_inputs"].items():
        verify_file(metadata, label)
        checked += 1
    for metadata in summary["capture"]["selected_dump_artifacts"]:
        verify_file(metadata, f"selected dump {metadata['path']}")
        checked += 1
    for label, metadata in summary["artifact_reports"].items():
        verify_file(metadata, label)
        checked += 1

    source = Path(summary["resolver_clone_site"]["source_file"])
    source_lines = source.read_text().splitlines()
    line_number = summary["resolver_clone_site"]["line"]
    check(source_lines[line_number - 1].strip() == summary["resolver_clone_site"]["line_text"],
          "source clone line changed")

    owners = [verify_owner(owner) for owner in summary["owner_proof"]]
    repeats = [
        verify_repeat(summary["repeat_owner_values"]["query_snapshot"], 16),
        verify_repeat(summary["repeat_owner_values"]["publication_current_snapshot_subcall"], 17),
    ]
    return {
        "status": "verified",
        "artifacts_checked": checked,
        "owners": owners,
        "repeats": repeats,
        "claim": summary["claim"],
    }


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--packet", type=Path, default=Path(__file__).resolve().parent)
    args = parser.parse_args()
    try:
        print(json.dumps(replay(args.packet), sort_keys=True))
    except (OSError, KeyError, TypeError, json.JSONDecodeError, ProfileError, AttributeError) as error:
        print(f"profile replay failed: {error}", file=sys.stderr)
        sys.exit(1)
