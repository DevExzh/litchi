#!/usr/bin/env python3
"""Verify reviewed source, compile-first gates, and generated XML receipts."""

import hashlib
import json
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = next(path for path in HERE.parents if (path / "crates").is_dir())


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main():
    gates = json.loads((HERE / "gates/receipt.json").read_text())
    assert gates["passed"] and gates["source_unchanged"]
    assert [row["name"] for row in gates["commands"]] == [
        "compile", "tests", "clippy", "rustdoc", "format", "diff"
    ]
    for row in gates["commands"]:
        assert row["exit_code"] == 0
        assert not row["RUSTFLAGS"] and not row["CARGO_ENCODED_RUSTFLAGS"]
        assert not row["RUSTC_BOOTSTRAP"]
    sources = json.loads((HERE / "gates/source-hashes.json").read_text())
    for name, expected in sources.items():
        assert digest(ROOT / name) == expected, name
    for name, expected in gates["logs"].items():
        assert digest(HERE / "gates" / name) == expected, name
    review = json.loads((HERE / "review.json").read_text())
    assert review["verdict"] == "approved"
    for name, expected in review["source_hashes"].items():
        assert sources[name] == expected, name
    probe = json.loads((HERE / "probe-receipt.json").read_text())
    assert probe["passed"] and probe["source_unchanged"]
    assert len(probe["commands"]) == 2
    assert all(row["exit_code"] == 0 for row in probe["commands"])
    for group in ("inputs", "outputs"):
        for name, expected in probe[group].items():
            assert digest(ROOT / name) == expected, name
    schema = json.loads((HERE / "schema-validation.json").read_text())
    assert len(schema["reports"]) == 3
    for row in schema["reports"]:
        assert row["valid"] and digest(ROOT / row["path"]) == row["sha256"], row["path"]
    result = {
        "passed": True,
        "source_files_checked": len(sources),
        "reviewed_files": len(review["source_hashes"]),
        "test_totals_including_doctests": gates["test_totals_including_doctests"],
        "schema_valid_elements": len(schema["reports"]),
        "gates_receipt_sha256": digest(HERE / "gates/receipt.json"),
        "probe_receipt_sha256": digest(HERE / "probe-receipt.json"),
    }
    (HERE / "verification.json").write_text(json.dumps(result, indent=2) + "\n")
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
