#!/usr/bin/env python3
"""Exercise fail-closed checks in the 0904 XML census helper.

This is a temporary-fixture self-test.  It imports ``census.py`` directly,
creates no build or benchmark artifacts, and removes every fixture when each
case exits.  The root coordinator runs it only after freezing the corpus and
the census helper.
"""

from __future__ import annotations

import copy
import hashlib
import importlib.util
import json
from pathlib import Path
import sys
import tempfile
import warnings
import zipfile


HERE = Path(__file__).resolve().parent
MEMBER = "xl/workbook.xml"
ORIGINAL = b'<root xmlns:p="urn:test"><p:item a="1"/><empty/></root>'
REPLACEMENT = b'<root xmlns:p="urn:test"><p:item a="2"/><empty/></root>'
COUNTS = {
    "start_or_empty_tags": 3,
    "zero_attribute_tags": 1,
    "one_attribute_tags": 2,
    "multiple_attribute_tags": 0,
    "attributes": 2,
    "xml_space_attributes": 0,
    "xml_space_elements": 0,
}


def _load_census():
    spec = importlib.util.spec_from_file_location("litchi_0904_census", HERE / "census.py")
    if spec is None or spec.loader is None:
        raise RuntimeError("could not load census.py")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _descriptor(path: str, data: bytes) -> dict[str, object]:
    return {"path": path, "bytes": len(data), "sha256": _sha256(data)}


def _write_zip(path: Path, entries: list[tuple[str, bytes]]) -> bytes:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_STORED) as archive:
            for name, data in entries:
                archive.writestr(name, data)
    return path.read_bytes()


def _manifest(
    root: Path,
    original: bytes = ORIGINAL,
    replacement: bytes = REPLACEMENT,
    source_entries: list[tuple[str, bytes]] | None = None,
    output_entries: list[tuple[str, bytes]] | None = None,
    proof: dict[str, object] | None = None,
) -> Path:
    original_path = root / "original.xml"
    replacement_path = root / "replacement.xml"
    original_path.write_bytes(original)
    replacement_path.write_bytes(replacement)
    source_entries = source_entries if source_entries is not None else [(MEMBER, original)]
    output_entries = output_entries if output_entries is not None else [(MEMBER, replacement)]
    source_path = root / "source.zip"
    output_path = root / "output.zip"
    source_bytes = _write_zip(source_path, source_entries)
    output_bytes = _write_zip(output_path, output_entries)
    if proof is None:
        proof = {"kind": "Complete"}
    manifest = {
        "schema": "litchi.xlsx-attribute-fixtures.v1",
        "cases": [
            {
                "id": "test",
                "shape": "test",
                "case": "mutation",
                "source": _descriptor("source.zip", source_bytes),
                "output": _descriptor("output.zip", output_bytes),
                "parts": [
                    {
                        "member": MEMBER,
                        "original": _descriptor("original.xml", original),
                        "replacement": _descriptor("replacement.xml", replacement),
                        "original_census": copy.deepcopy(COUNTS),
                        "replacement_census": copy.deepcopy(COUNTS),
                        "proof": proof,
                    }
                ],
            }
        ],
    }
    manifest_path = root / "manifest.json"
    manifest_path.write_text(json.dumps(manifest, sort_keys=True), encoding="utf-8")
    return manifest_path


def _expect_failure(census, root: Path, label: str) -> None:
    try:
        census.build_result(root, root / "manifest.json")
    except census.CensusError:
        return
    raise AssertionError(f"{label}: census unexpectedly accepted the mutation")


def _case(census, label: str, mutate) -> None:
    with tempfile.TemporaryDirectory(prefix="litchi-0904-census-test-") as directory:
        root = Path(directory)
        _manifest(root)
        mutate(root)
        _expect_failure(census, root, label)


def main() -> int:
    census = _load_census()

    with tempfile.TemporaryDirectory(prefix="litchi-0904-census-test-") as directory:
        root = Path(directory)
        _manifest(root)
        result = census.build_result(root, root / "manifest.json")
        assert result["totals"]["case_count"] == 1
        assert result["totals"]["part_count"] == 1

    def mismatched_zip_hash(root: Path) -> None:
        manifest = json.loads((root / "manifest.json").read_text(encoding="utf-8"))
        manifest["cases"][0]["source"]["sha256"] = "0" * 64
        (root / "manifest.json").write_text(json.dumps(manifest), encoding="utf-8")

    _case(census, "mismatched ZIP hash", mismatched_zip_hash)

    def sidecar_member_mismatch(root: Path) -> None:
        replacement = b'<root xmlns:p="urn:test"><p:item a="9"/><empty/></root>'
        (root / "original.xml").write_bytes(replacement)
        manifest = json.loads((root / "manifest.json").read_text(encoding="utf-8"))
        manifest["cases"][0]["parts"][0]["original"] = _descriptor(
            "original.xml", replacement
        )
        (root / "manifest.json").write_text(json.dumps(manifest), encoding="utf-8")

    _case(census, "sidecar/member mismatch", sidecar_member_mismatch)

    def missing_selected_member(root: Path) -> None:
        source_bytes = _write_zip(root / "source.zip", [("xl/other.xml", ORIGINAL)])
        manifest = json.loads((root / "manifest.json").read_text(encoding="utf-8"))
        manifest["cases"][0]["source"] = _descriptor("source.zip", source_bytes)
        (root / "manifest.json").write_text(json.dumps(manifest), encoding="utf-8")

    _case(census, "missing selected member", missing_selected_member)

    def duplicate_selected_member(root: Path) -> None:
        source_bytes = _write_zip(root / "source.zip", [(MEMBER, ORIGINAL), (MEMBER, ORIGINAL)])
        manifest = json.loads((root / "manifest.json").read_text(encoding="utf-8"))
        manifest["cases"][0]["source"] = _descriptor("source.zip", source_bytes)
        (root / "manifest.json").write_text(json.dumps(manifest), encoding="utf-8")

    _case(census, "duplicate selected member", duplicate_selected_member)

    def false_window_span(root: Path) -> None:
        manifest = json.loads((root / "manifest.json").read_text(encoding="utf-8"))
        manifest["cases"][0]["parts"][0]["proof"] = {
            "kind": "Window",
            "original": [0, len(ORIGINAL)],
            "replacement": [0, len(REPLACEMENT) + 1],
        }
        (root / "manifest.json").write_text(json.dumps(manifest), encoding="utf-8")

    _case(census, "false proof span", false_window_span)

    dtd = b'<!DOCTYPE root [<!ENTITY bad "x">]><root>&bad;</root>'

    def dtd_input(root: Path) -> None:
        _manifest(root, original=dtd, replacement=dtd, proof={"kind": "Identical"})

    _case(census, "DTD input", dtd_input)

    external = b'<!DOCTYPE root SYSTEM "https://example.invalid/entity.dtd"><root/>'

    def external_entity_input(root: Path) -> None:
        _manifest(root, original=external, replacement=external, proof={"kind": "Identical"})

    _case(census, "external entity input", external_entity_input)
    print("0904 census mutation checks: PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
