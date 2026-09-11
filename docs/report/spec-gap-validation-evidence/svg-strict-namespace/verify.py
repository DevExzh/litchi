#!/usr/bin/env python3
"""Recompute the compile-first and downstream receipts."""

import hashlib
import json
from pathlib import Path


HERE = Path(__file__).resolve().parent
ROOT = next(path for path in HERE.parents if (path / "crates").is_dir())


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main():
    compile_receipt = json.loads((HERE / "gates/receipt.json").read_text())
    assert compile_receipt["passed"] and compile_receipt["source_unchanged"]
    manifest = json.loads((HERE / "gates/source-hashes.json").read_text())
    for name, expected in manifest.items():
        assert digest(ROOT / name) == expected, name
    assert digest(HERE / "gates/source-hashes.json") == compile_receipt["source_hashes_sha256"]
    assert digest(HERE / "gates/compile-first.log") == compile_receipt["compile_log_sha256"]

    probe = json.loads((HERE / "probe-receipt.json").read_text())
    assert probe["passed"] and probe["source_unchanged"]
    assert digest(HERE / "gates/receipt.json") == probe["compile_first_receipt_sha256"]
    for name, expected in probe["generated_outputs"].items():
        assert digest(ROOT / name) == expected, name
    for name, expected in probe["raw_results"].items():
        assert digest(ROOT / name) == expected, name
    schema = json.loads((HERE / "raw-results/schema-validation.json").read_text())
    assert schema["passed"]
    slides = [row for row in schema["reports"] if row["kind"] == "strict-presentation-slide"]
    assert [len(row["svg_owners"]) for row in slides] == [0, 1, 0]
    assert all(
        value.startswith("http://purl.oclc.org/ooxml/officeDocument/relationships/")
        for row in slides
        for value in row["physical_relationship_types"]
    )
    controls = [row for row in schema["reports"] if row["kind"] == "direct-ms-odrawxml-svg-child"]
    assert [row["schema_valid"] for row in controls] == [True, False]
    xlsx = [
        row for row in schema["reports"]
        if row["kind"] == "strict-native-xlsx-mixed-namespace-drawing"
    ]
    assert len(xlsx) == 1 and len(xlsx[0]["svg_owners"]) == 2
    result = {
        "passed": True,
        "compile_first_receipt_sha256": digest(HERE / "gates/receipt.json"),
        "probe_receipt_sha256": digest(HERE / "probe-receipt.json"),
        "schema_validation_sha256": digest(HERE / "raw-results/schema-validation.json"),
        "source_files_checked": len(manifest),
        "generated_outputs": len(probe["generated_outputs"]),
        "scope": "Recomputed bounded Strict/Transitional namespace receipts; no final feature approval",
    }
    (HERE / "verification.json").write_text(json.dumps(result, indent=2) + "\n")
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
