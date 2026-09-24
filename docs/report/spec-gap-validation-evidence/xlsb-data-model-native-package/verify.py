#!/usr/bin/env python3
"""Verify the source-bound native-package integration evidence."""
import hashlib
import json
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = next(p for p in HERE.parents if (p / "crates").is_dir())

def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def main():
    record = json.loads((HERE / "validation.json").read_text())
    assert record["passed"] and record["source_unchanged"]
    assert not any(record["suppression_variables"].values())
    sources = json.loads((HERE / "source-hashes.json").read_text())
    assert len(sources) == record["source_files"]
    for path, expected in sources.items():
        assert digest(ROOT / path) == expected, path
    assert sources["test-data/ooxml/xlsb/date.xlsb"] == "fbb969989aabed2e057b2ddfc4cfd10ef1f6ebe9b7697c3755283515158d5b09"
    assert len(record["commands"]) == 2
    for command in record["commands"]:
        assert command["exit_code"] == 0
        assert digest(HERE / command["log"]) == command["sha256"]
    assert "2 passed; 0 failed; 0 ignored" in (HERE / "test.log").read_text()
    input_patch = record["preexisting_build_input"]
    assert digest(HERE / input_patch["reproduction_patch"]) == input_patch["sha256"]
    print(f"Verified {len(sources)} source inputs and native fixture; 2 lifecycle tests passed.")

if __name__ == "__main__":
    main()
