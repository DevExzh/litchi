#!/usr/bin/env python3
"""Read-only replay of the change-0817 admission-0 XML diagnosis.

This intentionally imports only Python's ZIP/XML readers.  It does not import
the frozen auditor and never starts a workload or subprocess.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path


CASES = {
    "generated-docx-medium": "docx",
    "generated-xlsx-medium": "xlsx",
    "generated-pptx-medium": "pptx",
    "real-000-docx": "docx",
    "real-001-xlsx": "xlsx",
    "real-002-pptx": "pptx",
}


def sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def archive(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as zf:
        return {name: zf.read(name) for name in zf.namelist()}


def canonical(data: bytes) -> bytes:
    root = ET.fromstring(data)
    raw = ET.tostring(root, encoding="utf-8", short_empty_elements=True)
    value = ET.canonicalize(xml_data=raw, with_comments=True)
    return value.encode("utf-8") if isinstance(value, str) else value


def attrs(data: bytes, local: str) -> dict[str, str] | None:
    root = ET.fromstring(data)
    for child in root:
        if child.tag.rsplit("}", 1)[-1] == local:
            return dict(child.attrib)
    return None


def relationship_rows(data: bytes) -> list[tuple[tuple[str, str], ...]]:
    root = ET.fromstring(data)
    return [tuple(sorted(child.attrib.items())) for child in root]


def namespace_census(data: bytes) -> list[list[str]]:
    import io

    return [[prefix or "", uri] for _event, (prefix, uri) in ET.iterparse(
        io.BytesIO(data), events=("start-ns",)
    )]


def member_changes(source: dict[str, bytes], output: dict[str, bytes]) -> list[str]:
    return sorted(
        set(source) ^ set(output)
        | {name for name in set(source) & set(output) if source[name] != output[name]}
    )


def policy_paths(case_dir: Path, extension: str) -> dict[str, Path]:
    return {
        policy: case_dir / f"{policy}.{extension}"
        for policy in ("default", "full", "file-only", "no-sync", "stream")
    }


def case_report(root: Path, manifest_case: dict) -> dict:
    case_id = manifest_case["case_id"]
    extension = CASES[case_id]
    case_dir = root / case_id
    source_path = case_dir / f"source.{extension}"
    source = archive(source_path)
    outputs = {policy: archive(path) for policy, path in policy_paths(case_dir, extension).items()}
    changed = {policy: member_changes(source, output) for policy, output in outputs.items()}
    output_hashes = {
        policy: sha(path.read_bytes())
        for policy, path in policy_paths(case_dir, extension).items()
    }
    result = {
        "case_id": case_id,
        "source_members": len(source),
        "source_decompressed_bytes": sum(map(len, source.values())),
        "output_hashes": output_hashes,
        "all_policy_outputs_equal": len(set(output_hashes.values())) == 1,
        "changed_members": changed,
    }
    corpus = manifest_case["corpus"]
    if manifest_case["origin"] == "generated-harness-corpus":
        result["generated_declared_vs_physical"] = {
            "entry_count": {"declared": corpus["entry_count"], "physical_members": len(source)},
            "archive_member_count": {
                "declared": corpus["archive_member_count"],
                "physical_members": len(source),
            },
            "uncompressed_payload_bytes": {
                "declared": corpus["uncompressed_payload_bytes"],
                "physical_bytes": sum(map(len, source.values())),
            },
            "target_entry": corpus["target_entry"],
            "target_payload_bytes": corpus["target_payload_bytes"],
        }
    if case_id == "real-000-docx":
        name = "word/_rels/document.xml.rels"
        source_rows = relationship_rows(source[name])
        output_rows = relationship_rows(outputs["default"][name])
        result["relationship_diagnosis"] = {
            "source_order": [dict(row)["Id"] for row in source_rows],
            "output_order": [dict(row)["Id"] for row in output_rows],
            "edge_set_equal": set(source_rows) == set(output_rows),
            "raw_equal": source[name] == outputs["default"][name],
            "canonical_equal": canonical(source[name]) == canonical(outputs["default"][name]),
            "namespace_census_equal": namespace_census(source[name])
            == namespace_census(outputs["default"][name]),
            "source_bytes": len(source[name]),
            "output_bytes": len(outputs["default"][name]),
        }
    if case_id in {"real-001-xlsx", "generated-xlsx-medium"}:
        name = "xl/workbook.xml"
        source_root = ET.fromstring(source[name])
        output_root = ET.fromstring(outputs["default"][name])
        source_without_calc = [child for child in source_root if child.tag.rsplit("}", 1)[-1] != "calcPr"]
        output_without_calc = [child for child in output_root if child.tag.rsplit("}", 1)[-1] != "calcPr"]
        result["workbook_diagnosis"] = {
            "source_calcPr": attrs(source[name], "calcPr"),
            "output_calcPr": attrs(outputs["default"][name], "calcPr"),
            "structure_without_calcPr_equal": [ET.tostring(x) for x in source_without_calc]
            == [ET.tostring(x) for x in output_without_calc],
            "raw_equal": source[name] == outputs["default"][name],
            "canonical_equal": canonical(source[name]) == canonical(outputs["default"][name]),
            "namespace_census_equal": namespace_census(source[name])
            == namespace_census(outputs["default"][name]),
            "source_bytes": len(source[name]),
            "output_bytes": len(outputs["default"][name]),
        }
    return result


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("artifacts", type=Path)
    args = parser.parse_args()
    manifest_path = args.artifacts / "manifest.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    report_path = args.artifacts.parent / "admission-0" / "artifact-audit.json"
    report = json.loads(report_path.read_text(encoding="utf-8"))
    result = {
        "report_ok": report.get("ok"),
        "report_error_count": len(report.get("errors", [])),
        "report_errors": report.get("errors", []),
        "cases": [case_report(args.artifacts, case) for case in manifest["cases"]],
    }
    print(json.dumps(result, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
