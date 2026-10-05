#!/usr/bin/env python3
"""Replay the retained 0916 DOCX Callgrind caller proof.

The script only reads the small packet in this directory and the raw external
artifacts named by it. It never rebuilds the binary, writes to the external
profile, or overwrites a retained capture.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import re
import sys
import tarfile


class ProfileError(ValueError):
    """Raised when an external artifact does not match the packet."""


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


def callgrind_records(path: Path) -> tuple[str, list[dict]]:
    """Read caller edges and the first event value after each calls line."""

    text = path.read_text(errors="strict")
    lines = text.splitlines()
    function_names: dict[str, str] = {}
    for line in lines:
        match = re.match(r"fn=\((\d+)\)(?:\s+(.*))?$", line)
        if match is not None and match.group(2):
            function_names[match.group(1)] = match.group(2)

    records: list[dict] = []
    current_fn = ""
    pending_cfn: str | None = None
    for index, line in enumerate(lines):
        if line.startswith("fn="):
            match = re.match(r"fn=\((\d+)\)(?:\s+(.*))?$", line)
            current_fn = (
                match.group(2)
                if match is not None and match.group(2)
                else function_names.get(match.group(1), "")
                if match is not None
                else line
            )
            pending_cfn = None
            continue
        if line.startswith("cfn="):
            match = re.match(r"cfn=\((\d+)\)(?:\s+(.*))?$", line)
            pending_cfn = (
                match.group(2)
                if match is not None and match.group(2)
                else function_names.get(match.group(1), "")
                if match is not None
                else line
            )
            continue
        if pending_cfn is None or not line.startswith("calls="):
            continue
        match = re.match(r"calls=(\d+)", line)
        if match is None:
            continue
        cost = None
        for following in lines[index + 1 : index + 10]:
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


def header_value(text: str, name: str) -> str:
    match = re.search(rf"^{re.escape(name)}:\s*(.*)$", text, re.MULTILINE)
    check(match is not None, f"missing Callgrind header {name}")
    return match.group(1)


def verify_owner(owner: dict, expected_trigger: str) -> dict:
    dump = verify_file(
        {"path": owner["dump"], "bytes": owner["dump_bytes"], "sha256": owner["dump_sha256"]},
        owner["owner"],
    )
    text, records = callgrind_records(dump)
    check(int(header_value(text, "part")) == owner["part"], f"{owner['owner']}: part mismatch")
    check(header_value(text, "events") == "Ir", f"{owner['owner']}: event mismatch")
    check(expected_trigger in text, f"{owner['owner']}: trigger mismatch")
    root_prefix = owner["root_prefix"]
    root = one_record(
        records,
        lambda record: record["callee"].startswith(root_prefix)
        and all(fragment in record["callee"] for fragment in owner["root_fragments"])
        and record["calls"] == 0,
        f"{owner['owner']}: root caller chain",
    )
    # The packet stores the selected owner's total once as summary_ir.  Read
    # that value for the root record so the manifest does not duplicate a
    # hand-entered counter under a second key.
    check(root["cost"] == owner["summary_ir"], f"{owner['owner']}: root Ir mismatch")

    nested = {}
    for key in ("patch_apply", "snapshot", "ensure", "parse"):
        expected = owner[key]
        record = one_record(
            records,
            lambda candidate, expected=expected: candidate["callee"].startswith(expected["prefix"])
            and all(fragment in candidate["callee"] for fragment in expected["fragments"])
            and candidate["calls"] == expected["calls"],
            f"{owner['owner']}: {key} caller",
        )
        check(record["cost"] == expected["ir"], f"{owner['owner']}: {key} Ir mismatch")
        nested[key] = {"calls": record["calls"], "ir": record["cost"]}
    return {
        "owner": owner["owner"],
        "part": owner["part"],
        "root_ir": root["cost"],
        "nested": nested,
    }


def verify_report(path: Path, binary: dict, corpus: dict) -> None:
    report = json.loads(path.read_text())
    identity = report["binary_identity"]
    check(identity["binary_sha256"] == binary["sha256"], f"{path}: binary hash mismatch")
    check(identity["binary_bytes"] == binary["bytes"], f"{path}: binary size mismatch")
    result = report["results"]
    check(len(result) == 1, f"{path}: expected one result")
    case = result[0]
    check(case["case"] == "docx_source_backed_existing_section_layout_edit_save", f"{path}: case mismatch")
    check(case["corpus"]["archive_sha256"] == corpus["archive_sha256"], f"{path}: corpus hash mismatch")
    check(case["corpus"]["target_payload_sha256"] == corpus["target_payload_sha256"], f"{path}: target hash mismatch")


def verify_source_member(summary: dict) -> None:
    archive = Path(summary["frozen_inputs"]["source_archive"]["path"])
    member = summary["frozen_inputs"]["source_member_line_proof"]
    check(archive.is_file(), f"source archive missing: {archive}")
    with tarfile.open(archive, "r:") as source_tar:
        entry = source_tar.extractfile(member["archive_member"])
        check(entry is not None, f"source archive member missing: {member['archive_member']}")
        data = entry.read()
    check(len(data) == member["bytes"], "source member size mismatch")
    check(hashlib.sha256(data).hexdigest() == member["sha256"], "source member hash mismatch")
    lines = data.decode().splitlines()
    check(lines[member["apply_line"] - 1].strip() == member["apply_line_text"], "apply line proof mismatch")
    check(lines[member["target_parse_line"] - 1].strip() == member["target_parse_line_text"], "target parse line proof mismatch")


def replay(packet: Path) -> dict:
    summary_path = packet / "profile-reduced.json"
    summary = json.loads(summary_path.read_text())
    binary = summary["frozen_inputs"]["binary"]
    source_archive = summary["frozen_inputs"]["source_archive"]
    verify_file(binary, "binary")
    verify_file(source_archive, "source archive")
    verify_source_member(summary)

    corpus = summary["corpus"]
    report_paths = [
        next(item for item in summary["external_artifacts"] if item["path"].endswith("/raw/patch-apply.json")),
        next(item for item in summary["external_artifacts"] if item["path"].endswith("/raw/publication.json")),
    ]
    for report_metadata in report_paths:
        report = verify_file(report_metadata, "raw harness report")
        verify_report(report, binary, corpus)

    checked = 2 + len(report_paths)
    for metadata in summary["external_artifacts"]:
        verify_file(metadata, "external artifact")
        checked += 1

    target = summary["owner_proof"]["target_patch_apply"]
    publication = summary["owner_proof"]["whole_publication"]
    target["dump_bytes"] = next(item["bytes"] for item in summary["external_artifacts"] if item["path"] == target["dump"])
    target["dump_sha256"] = next(item["sha256"] for item in summary["external_artifacts"] if item["path"] == target["dump"])
    publication["dump_bytes"] = next(item["bytes"] for item in summary["external_artifacts"] if item["path"] == publication["dump"])
    publication["dump_sha256"] = next(item["sha256"] for item in summary["external_artifacts"] if item["path"] == publication["dump"])
    target_result = verify_owner(
        target | {
            "owner": "target_patch_apply",
            "root_prefix": "litchi_docx::section::layout::Patch::apply",
            "root_fragments": ["publish_section_layout_commit_to_stream", "run_docx_source_backed_existing_section_layout_edit_save"],
            "patch_apply": {"prefix": "litchi_docx::section::layout::Patch::apply", "fragments": ["publish_section_layout_commit_to_stream"], "calls": 0, "ir": target["patch_apply_ir"]},
            "snapshot": {"prefix": "litchi_docx::section::layout::Snapshot::from_shared_xml", "fragments": ["Patch::apply"], "calls": target["snapshot_call_count"], "ir": target["snapshot_from_shared_xml_ir"]},
            "ensure": {"prefix": "litchi_docx::source_backed::ensure_source_section_inventory_xml", "fragments": ["Snapshot::from_shared_xml", "Patch::apply"], "calls": target["ensure_call_count"], "ir": target["ensure_source_section_inventory_xml_ir"]},
            "parse": {"prefix": "litchi_docx::section::inventory::parse_inventory", "fragments": ["Snapshot::from_shared_xml", "Patch::apply"], "calls": target["parse_call_count"], "ir": target["parse_inventory_ir"]},
        },
        "litchi_docx::section::layout::Patch::apply",
    )
    publication_result = verify_owner(
        publication | {
            "owner": "whole_publication",
            "root_prefix": "litchi_docx::source_backed::Package::publish_section_layout_commit_to_stream",
            "root_fragments": ["run_docx_source_backed_existing_section_layout_edit_save"],
            "patch_apply": {"prefix": "litchi_docx::section::layout::Patch::apply", "fragments": ["publish_section_layout_commit_to_stream"], "calls": publication["patch_apply_call_count"], "ir": publication["patch_apply_ir"]},
            "snapshot": {"prefix": "litchi_docx::section::layout::Snapshot::from_shared_xml", "fragments": ["Patch::apply"], "calls": publication["snapshot_call_count"], "ir": publication["snapshot_from_shared_xml_ir"]},
            "ensure": {"prefix": "litchi_docx::source_backed::ensure_source_section_inventory_xml", "fragments": ["Snapshot::from_shared_xml", "Patch::apply"], "calls": publication["ensure_call_count"], "ir": publication["ensure_source_section_inventory_xml_ir"]},
            "parse": {"prefix": "litchi_docx::section::inventory::parse_inventory", "fragments": ["Snapshot::from_shared_xml", "Patch::apply"], "calls": publication["parse_call_count"], "ir": publication["parse_inventory_ir"]},
        },
        "litchi_docx::source_backed::Package::publish_section_layout_commit_to_stream",
    )
    expected_share = summary["derived"]["target_parse_share_of_whole_publication_percent"]
    actual_share = publication["snapshot_from_shared_xml_ir"] / publication["summary_ir"] * 100
    check(abs(actual_share - expected_share) < 0.000001, "derived publication share mismatch")
    return {
        "status": "verified",
        "artifacts_checked": checked,
        "owners": [target_result, publication_result],
        "claim": summary["claim"],
    }


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--packet", type=Path, default=Path(__file__).resolve().parent)
    args = parser.parse_args()
    try:
        print(json.dumps(replay(args.packet), sort_keys=True))
    except (OSError, KeyError, TypeError, ValueError, json.JSONDecodeError, ProfileError, tarfile.TarError) as error:
        print(f"profile replay failed: {error}", file=sys.stderr)
        sys.exit(1)
