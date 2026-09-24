#!/usr/bin/env python3
"""Verify the compile-first closure receipts against their reviewed inputs."""
import hashlib
import json
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = next(p for p in HERE.parents if (p / "crates").is_dir())

def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

receipt = json.loads((HERE / "gates/receipt.json").read_text())
assert receipt["passed"] and receipt["source_unchanged"]
assert [row["name"] for row in receipt["commands"]] == [
    "compile", "tests", "clippy", "rustdoc", "format", "diff"]
assert all(row["exit_code"] == 0 for row in receipt["commands"])
sources = json.loads((HERE / "gates/source-hashes.json").read_text())
for name, expected in sources.items():
    assert digest(ROOT / name) == expected, name
for name, expected in receipt["logs"].items():
    assert digest(HERE / "gates" / name) == expected, name
review = json.loads((HERE / "review.json").read_text())
assert review["verdict"] == "approved"
for name, expected in review["source_hashes"].items():
    assert sources[name] == expected, name
result = {
    "passed": True,
    "compile_first": True,
    "source_files_checked": len(sources),
    "reviewed_production_files": len(review["source_hashes"]),
    "test_totals_including_doctests": receipt["test_totals_including_doctests"],
    "receipt_sha256": digest(HERE / "gates/receipt.json"),
}
(HERE / "verification.json").write_text(json.dumps(result, indent=2) + "\n")
print(json.dumps(result, indent=2))
