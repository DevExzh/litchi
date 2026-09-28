"""Independently replay ordinary-save ZIP preservation evidence.

This checker is deliberately separate from the Rust exporter and from the
semantic XML auditor.  It checks every retained output for all five
publication policies, binds the three caller-named inputs and the six exact
historical output identities, and compares untouched ZIP members at the raw
compressed-payload and central-directory metadata level.  It never runs a
format reader or a workload.

The normal commands are::

    python3 -B preservation.py --leg before --write
    python3 -B preservation.py --leg after --write
    python3 -B preservation.py --leg before --check
    python3 -B preservation.py --leg after --check

The write mode refuses to overwrite its report.  Check mode recomputes the
report and requires byte-for-byte equality with the retained JSON, which is
the offline replay used by ``admission.py`` and the final validator.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import posixpath
import struct
import zipfile
from pathlib import Path
from typing import Any


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
HISTORICAL = ROOT / "docs/performance/results/change-0821/artifacts"
HISTORICAL_PRIOR = ROOT / "docs/performance/results/change-0819/artifacts"

POLICIES = ("default", "full", "file-only", "no-sync", "stream")
FORMATS = ("DOCX", "XLSX", "PPTX")

# These are the three exact source inputs sealed by the 0819/0821 admission
# records.  They are repeated here instead of being inferred from a current
# manifest so a changed input cannot silently become the new baseline.
INPUTS: dict[str, dict[str, Any]] = {
    "real-000-docx": {
        "format": "DOCX",
        "origin": "caller-named-real-file",
        "input": "test-data/ooxml/docx/documentProperties.docx",
        "bytes": 23_503,
        "sha256": "1cff7a0a94dfce307a70032d21070d26ae34b9fdf742cf70fa66d4a2078ec9d5",
        "reference": "real-000-docx/default.docx",
        "main_part": "word/document.xml",
        "target": {"main_part": "word/document.xml"},
        "closure": ("word/document.xml",),
    },
    "real-001-xlsx": {
        "format": "XLSX",
        "origin": "caller-named-real-file",
        "input": "test-data/libreoffice-core/sc/qa/unit/data/xlsx/dateAutofilter.xlsx",
        "bytes": 8_435,
        "sha256": "d7ab3dbb59388d245ee779bf8547748dc6bac70f3c7216e673e0d97dbbbd6bc4",
        "reference": "real-001-xlsx/default.xlsx",
        "main_part": "xl/workbook.xml",
        "target": {
            "main_part": "xl/workbook.xml",
            "xlsx_sheet": "Munka1",
            "xlsx_address": "A1",
        },
        "closure": ("xl/workbook.xml", "xl/worksheets/sheet1.xml"),
    },
    "real-002-pptx": {
        "format": "PPTX",
        "origin": "caller-named-real-file",
        "input": "test-data/ooxml/pptx/shapes.pptx",
        "bytes": 68_822,
        "sha256": "19fde9b87e33dd1a95fdbba0cf6abc2278bf03874f4665c7f8b88b6afe4a2571",
        "reference": "real-002-pptx/default.pptx",
        "main_part": "ppt/presentation.xml",
        "target": {
            "main_part": "ppt/presentation.xml",
            "pptx_slide": 0,
            "pptx_shape": 0,
        },
        "closure": ("ppt/slides/slide1.xml",),
    },
    "generated-docx-medium": {
        "format": "DOCX",
        "origin": "generated-harness-corpus",
        "bytes": 9_051,
        "sha256": "2004b8046b78320f2d05f27bce1ef293c1783912e2f27c83674e8f7d4c9431d9",
        "reference": "generated-docx-medium/default.docx",
        "main_part": "word/document.xml",
        "target": {"main_part": "word/document.xml"},
        "closure": ("word/document.xml",),
    },
    "generated-xlsx-medium": {
        "format": "XLSX",
        "origin": "generated-harness-corpus",
        "bytes": 4_226_429,
        "sha256": "dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036",
        "reference": "generated-xlsx-medium/default.xlsx",
        "main_part": "xl/workbook.xml",
        "target": {
            "main_part": "xl/workbook.xml",
            "xlsx_sheet": "Sheet1",
            "xlsx_address": "A1",
        },
        "closure": ("xl/workbook.xml", "xl/worksheets/sheet1.xml"),
    },
    "generated-pptx-medium": {
        "format": "PPTX",
        "origin": "generated-harness-corpus",
        "bytes": 40_788,
        "sha256": "50ad2f81099ee29d4768d7080b7fc51ea2b5ca2aadd531031efcab65e8d5409e",
        "reference": "generated-pptx-medium/default.pptx",
        "main_part": "ppt/presentation.xml",
        "target": {
            "main_part": "ppt/presentation.xml",
            "pptx_slide": 0,
            "pptx_shape": 0,
        },
        "closure": ("ppt/slides/slide1.xml",),
    },
}

# Published bytes are intentionally listed independently from the historical
# manifests.  The two historical artifact trees are then required to agree
# with these values and with each other.
PUBLISHED: dict[str, tuple[int, str]] = {
    "generated-docx-medium": (
        9_070,
        "6e2a7afe4d670787c7d428032a8fc82ca866ab9e17ba50523ad108c366bf3775",
    ),
    "generated-xlsx-medium": (
        4_226_568,
        "20335c4480405e051be2f4fea0e2a5aa8c5912b729cb1df980112ea7c5056687",
    ),
    "generated-pptx-medium": (
        40_802,
        "95efb7b3b4f621b68961bb3db87618cbe59065650a11b64ead3debc8e598d2cd",
    ),
    "real-000-docx": (
        23_535,
        "de9e163ac26e170ee3881c7d7e53ac29836efd6caea6f72db7ebd2cd31d6c774",
    ),
    "real-001-xlsx": (
        8_521,
        "0f6902152f1b8c40c47086206023887e87f54ef57f1da417bc91d8494a4d3e68",
    ),
    "real-002-pptx": (
        68_284,
        "38c7fd3cc5037316e0a3a9c8cb6e6b013a95dd03b39bbd7fb357fc3cc3b9e5cf",
    ),
}

METADATA_FIELDS = (
    "date_time",
    "compress_type",
    "flag_bits",
    "create_system",
    "create_version",
    "extract_version",
    "reserved",
    "volume",
    "internal_attr",
    "external_attr",
    "extra",
    "comment",
    "CRC",
    "file_size",
    "compress_size",
)


def digest(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def descriptor(path: Path) -> dict[str, Any]:
    assert path.is_file() and not path.is_symlink(), f"missing or symlink artifact: {path}"
    data = path.read_bytes()
    return {"path": str(path), "bytes": len(data), "sha256": digest(data)}


def json_read(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def safe_relative(value: str, root: Path) -> Path:
    candidate = Path(value)
    assert not candidate.is_absolute(), f"absolute artifact path: {value!r}"
    assert "\\" not in value, f"backslash artifact path: {value!r}"
    resolved = (root / candidate).resolve()
    root_resolved = root.resolve()
    assert resolved == root_resolved or root_resolved in resolved.parents, (
        f"artifact path escapes output directory: {value!r}"
    )
    assert str(candidate) == posixpath.normpath(str(candidate)), (
        f"non-normalized artifact path: {value!r}"
    )
    return resolved


def zip_member_name_safe(name: str) -> bool:
    if not name or name.startswith(("/", "\\")) or "\\" in name:
        return False
    if ":" in name.split("/", 1)[0]:
        return False
    parts = name.rstrip("/").split("/")
    return all(part not in ("", ".", "..") for part in parts)


def raw_compressed_payload(archive: bytes, info: zipfile.ZipInfo) -> bytes:
    offset = info.header_offset
    assert archive[offset : offset + 4] == b"PK\x03\x04", info.filename
    assert offset + 30 <= len(archive), info.filename
    name_length, extra_length = struct.unpack_from("<HH", archive, offset + 26)
    start = offset + 30 + name_length + extra_length
    end = start + info.compress_size
    assert end <= len(archive), info.filename
    return archive[start:end]


def zip_metadata(info: zipfile.ZipInfo) -> dict[str, Any]:
    return {
        field: (getattr(info, field).hex() if field == "extra" else getattr(info, field))
        for field in METADATA_FIELDS
    }


def read_zip(path: Path) -> dict[str, Any]:
    data = path.read_bytes()
    members: list[dict[str, Any]] = []
    seen: set[str] = set()
    with zipfile.ZipFile(path) as archive:
        assert archive.comment is not None
        for info in archive.infolist():
            name = info.filename
            assert zip_member_name_safe(name), f"unsafe ZIP member name {name!r}: {path}"
            assert name not in seen, f"duplicate ZIP member {name!r}: {path}"
            seen.add(name)
            decoded = archive.read(info)
            compressed = raw_compressed_payload(data, info)
            members.append(
                {
                    "name": name,
                    "data": decoded,
                    "compressed": compressed,
                    "metadata": zip_metadata(info),
                    "decoded_sha256": digest(decoded),
                    "compressed_sha256": digest(compressed),
                }
            )
        names = [item["name"] for item in members]
    return {
        "path": path,
        "bytes": data,
        "members": members,
        "by_name": {item["name"]: item for item in members},
        "names": names,
        "comment": archive.comment,
    }


def historical_reference(case_id: str) -> tuple[Path, Path]:
    expected = INPUTS[case_id]
    current = HISTORICAL / expected["reference"]
    prior = HISTORICAL_PRIOR / expected["reference"]
    assert current.is_file() and prior.is_file(), case_id
    current_bytes, prior_bytes = current.read_bytes(), prior.read_bytes()
    assert current_bytes == prior_bytes, f"0819/0821 reference differs for {case_id}"
    expected_bytes, expected_sha = PUBLISHED[case_id]
    assert len(current_bytes) == expected_bytes
    assert digest(current_bytes) == expected_sha
    return current, prior


def expected_case_ids() -> set[str]:
    return set(INPUTS)


def manifest_descriptor_path(manifest_case: dict[str, Any], key: str, root: Path) -> Path:
    value = manifest_case.get(key)
    assert isinstance(value, dict), f"missing {key} descriptor"
    path = value.get("path")
    assert isinstance(path, str), f"missing {key}.path"
    return safe_relative(path, root)


def policy_specs(case: dict[str, Any]) -> dict[str, dict[str, Any]]:
    filesystem = case.get("policy_outputs")
    assert isinstance(filesystem, list) and len(filesystem) == 4
    result: dict[str, dict[str, Any]] = {}
    for spec in filesystem:
        assert isinstance(spec, dict)
        policy = spec.get("policy")
        assert policy in POLICIES[:4] and policy not in result, policy
        result[policy] = spec
    stream = case.get("stream_output")
    assert isinstance(stream, dict) and stream.get("policy") == "stream"
    assert "stream" not in result
    result["stream"] = stream
    assert set(result) == set(POLICIES)
    return result


def all_files(root: Path) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for path in sorted(root.rglob("*")):
        if path.is_symlink():
            raise AssertionError(f"artifact tree contains symlink: {path}")
        if path.is_file():
            result.append(
                {
                    "path": str(path.relative_to(root)),
                    "bytes": path.stat().st_size,
                    "sha256": digest(path.read_bytes()),
                }
            )
    return result


def check_member_preservation(
    source: dict[str, Any],
    output: dict[str, Any],
    reference: dict[str, Any],
    expected_closure: tuple[str, ...],
) -> dict[str, Any]:
    assert source["names"] == output["names"] == reference["names"], "ZIP member order differs"
    assert source["comment"] == output["comment"] == reference["comment"], "ZIP archive comment differs"
    assert set(source["names"]) == set(reference["names"])
    closure = set(expected_closure)
    changed: list[str] = []
    untouched: list[dict[str, Any]] = []
    for name in source["names"]:
        left = source["by_name"][name]
        right = output["by_name"][name]
        historical = reference["by_name"][name]
        if right["data"] != left["data"]:
            changed.append(name)
            assert name in closure, f"member changed outside chosen edit closure: {name}"
            assert right["data"] == historical["data"], f"changed member differs from reference: {name}"
            continue
        assert name not in closure or historical["data"] == left["data"], name
        assert right["compressed"] == left["compressed"], f"untouched compressed payload changed: {name}"
        assert right["metadata"] == left["metadata"], f"untouched ZIP metadata changed: {name}"
        assert right["data"] == historical["data"], f"untouched member differs from reference: {name}"
        untouched.append(
            {
                "name": name,
                "decoded_sha256": digest(left["data"]),
                "compressed_sha256": digest(left["compressed"]),
                "metadata": left["metadata"],
                "metadata_equal": True,
                "compressed_payload_equal": True,
            }
        )
    assert tuple(sorted(changed)) == tuple(sorted(closure)), (
        f"changed member set {sorted(changed)} != chosen closure {sorted(closure)}"
    )
    return {
        "member_order_equal": True,
        "archive_comment_equal": True,
        "allowed_changed_members": sorted(closure),
        "changed_members": sorted(changed),
        "untouched_members": untouched,
    }


def analyze(leg: str) -> dict[str, Any]:
    artifacts = PACKET / f"artifacts-{leg}"
    assert artifacts.is_dir() and not artifacts.is_symlink(), artifacts
    manifest_path = artifacts / "manifest.json"
    assert manifest_path.is_file() and not manifest_path.is_symlink()
    manifest = json_read(manifest_path)
    assert manifest.get("schema_version") == 1
    assert manifest.get("kind") == "ordinary-save-artifact-export"
    assert manifest.get("generator") == "litchi-perf-ordinary-save-artifacts-v1"
    cases = manifest.get("cases")
    assert isinstance(cases, list) and len(cases) == 6
    by_id = {str(case.get("case_id")): case for case in cases}
    assert set(by_id) == expected_case_ids(), (set(by_id), expected_case_ids())

    case_reports: list[dict[str, Any]] = []
    expected_referenced: set[str] = {"manifest.json"}
    for case_id in sorted(expected_case_ids()):
        case = by_id[case_id]
        expected = INPUTS[case_id]
        assert case.get("format") == expected["format"]
        assert case.get("origin") == expected["origin"]
        assert case.get("edit_admitted") is True
        assert case.get("edit_outcome") == "admitted"
        assert case.get("refused_output_source_exact") is False
        assert case.get("edit_target") == expected["target"]

        source_path = manifest_descriptor_path(case, "source_archive", artifacts)
        source = read_zip(source_path)
        expected_referenced.add(str(source_path.relative_to(artifacts)))
        input_path = ROOT / expected["input"] if expected["origin"] == "caller-named-real-file" else None
        if input_path is not None:
            assert input_path.is_file() and not input_path.is_symlink()
            assert input_path.read_bytes() == source["bytes"]
        assert len(source["bytes"]) == expected["bytes"]
        assert digest(source["bytes"]) == expected["sha256"]
        assert case.get("source_archive_bytes") == len(source["bytes"])
        assert case.get("source_archive_sha256") == digest(source["bytes"])
        source_descriptor = case["source_archive"]
        assert source_descriptor.get("bytes") == len(source["bytes"])
        assert source_descriptor.get("sha256") == digest(source["bytes"])

        reference_path, prior_reference_path = historical_reference(case_id)
        reference = read_zip(reference_path)
        prior_reference = read_zip(prior_reference_path)
        assert reference["bytes"] == prior_reference["bytes"]
        expected_bytes, expected_sha = PUBLISHED[case_id]
        assert case.get("published_bytes") == expected_bytes
        assert case.get("published_sha256") == expected_sha

        specs = policy_specs(case)
        policy_reports: list[dict[str, Any]] = []
        default_output: dict[str, Any] | None = None
        for policy in POLICIES:
            spec = specs[policy]
            output_path = manifest_descriptor_path(spec, "output", artifacts)
            expected_referenced.add(str(output_path.relative_to(artifacts)))
            output = read_zip(output_path)
            output_bytes = output["bytes"]
            assert len(output_bytes) == expected_bytes
            assert digest(output_bytes) == expected_sha
            output_descriptor = spec.get("output")
            assert isinstance(output_descriptor, dict)
            assert output_descriptor.get("bytes") == expected_bytes
            assert output_descriptor.get("sha256") == expected_sha
            assert spec.get("bytes") == expected_bytes
            assert spec.get("sha256") == expected_sha
            assert spec.get("matches_reference") is True
            assert spec.get("matches_source") is False
            assert spec.get("source_unchanged") is True
            assert spec.get("reopen_admitted") is True
            if default_output is None:
                default_output = output
            else:
                assert output_bytes == default_output["bytes"], f"policy {policy} differs from default"
            preservation = check_member_preservation(
                source, output, reference, expected["closure"]
            )
            policy_reports.append(
                {
                    "policy": policy,
                    "publication": spec.get("publication"),
                    "path": str(output_path.relative_to(artifacts)),
                    "bytes": len(output_bytes),
                    "sha256": digest(output_bytes),
                    "matches_reference": True,
                    "matches_default": True,
                    **preservation,
                }
            )
        assert default_output is not None
        case_reports.append(
            {
                "case_id": case_id,
                "format": expected["format"],
                "origin": expected["origin"],
                "source": {
                    "path": str(source_path.relative_to(artifacts)),
                    "bytes": len(source["bytes"]),
                    "sha256": digest(source["bytes"]),
                },
                "historical_reference": {
                    "path": str(reference_path.relative_to(ROOT)),
                    "bytes": expected_bytes,
                    "sha256": expected_sha,
                    "0819_path": str(prior_reference_path.relative_to(ROOT)),
                    "0819_sha256": digest(prior_reference["bytes"]),
                },
                "allowed_changed_members": list(expected["closure"]),
                "policy_outputs": policy_reports,
            }
        )

    files = all_files(artifacts)
    actual_files = {row["path"] for row in files}
    assert actual_files == expected_referenced, (
        f"artifact tree is not fully bound; missing={sorted(expected_referenced - actual_files)}, "
        f"extra={sorted(actual_files - expected_referenced)}"
    )
    return {
        "schema": "litchi.performance.0825.zip-preservation.v1",
        "leg": leg,
        "artifact_directory": str(artifacts),
        "manifest": {
            "path": str(manifest_path),
            "bytes": manifest_path.stat().st_size,
            "sha256": digest(manifest_path.read_bytes()),
        },
        "cases": case_reports,
        "all_files": files,
        "checks": {
            "case_count": True,
            "five_policies_per_case": True,
            "source_input_identity": True,
            "historical_0819_0821_identity": True,
            "policy_outputs_byte_identical": True,
            "member_order": True,
            "archive_comments": True,
            "untouched_member_metadata": True,
            "untouched_compressed_payload": True,
            "chosen_edit_closures": True,
            "fresh_all_files_bound": True,
        },
        "ok": True,
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--leg", choices=("before", "after"), required=True)
    action = parser.add_mutually_exclusive_group()
    action.add_argument("--write", action="store_true")
    action.add_argument("--check", action="store_true")
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    target = args.output or PACKET / f"zip-preservation-{args.leg}.json"
    value = analyze(args.leg)
    encoded = json.dumps(value, indent=2, sort_keys=True) + "\n"
    if args.check:
        assert target.is_file(), target
        assert target.read_text(encoding="utf-8") == encoded, (
            f"retained preservation report does not replay exactly: {target}"
        )
        print(f"0825 preservation {args.leg} CHECK PASS")
        return 0
    assert args.write or not target.exists(), (
        "choose --write to create the report; refusing implicit overwrite"
    )
    assert not target.exists(), f"refusing to overwrite {target}"
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(encoded, encoding="utf-8")
    print(f"0825 preservation {args.leg} WRITE PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
