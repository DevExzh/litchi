#!/usr/bin/env python3
"""Recheck the retained test-only graph evidence against current source."""
import hashlib
import json
from pathlib import Path
import re

HERE = Path(__file__).resolve().parent
ROOT = next(p for p in HERE.parents if (p / "crates").is_dir())

def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def main():
    receipt = json.loads((HERE / "validation.json").read_text())
    assert receipt["passed"] and receipt["source_unchanged"]
    assert not any(receipt["suppression_variables"].values())
    sources = json.loads((HERE / "source-hashes.json").read_text())
    assert len(sources) == receipt["source_files"]
    for path, expected in sources.items():
        assert digest(ROOT / path) == expected, path
    assert len(receipt["commands"]) == 2
    for command in receipt["commands"]:
        assert command["exit_code"] == 0
        assert digest(HERE / command["log"]) == command["sha256"]
    totals = re.findall(r"test result: ok\. (\d+) passed; (\d+) failed; (\d+) ignored", (HERE / "tests.log").read_text())
    assert totals == [("17", "0", "0"), ("12", "0", "0")], totals
    test = "crates/litchi-xlsb/tests/data_model_connection_graph.rs"
    assert sources[test] == receipt["root_review"]["test_sha256"]
    print(f"Verified {len(sources)} source inputs; 29 tests passed (12 new graph tests).")

if __name__ == "__main__":
    main()
