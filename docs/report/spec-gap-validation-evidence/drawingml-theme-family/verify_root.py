#!/usr/bin/env python3
"""Independently check gate hashes and raw profile statistics/accounting."""

import hashlib
import json
import math
from pathlib import Path
import re

HERE = Path(__file__).resolve().parent
ROOT = next(p for p in HERE.parents if (p / "crates").is_dir())


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def quantile(values, percent):
    return sorted(values)[math.ceil(len(values) * percent / 100) - 1]


def main():
    gates = HERE / "gates"
    receipt = json.loads((gates / "receipt.json").read_text())
    assert receipt["passed"] and receipt["source_unchanged"]
    assert len(receipt["commands"]) == 5
    assert all(row["exit_code"] == 0 for row in receipt["commands"])
    source_hashes = json.loads((gates / "source-hashes.json").read_text())
    for name, expected in source_hashes.items():
        assert digest(ROOT / name) == expected, name
    for name, expected in receipt["logs"].items():
        assert digest(gates / name) == expected, name
    totals = re.findall(r"test result: ok\. (\d+) passed; (\d+) failed; (\d+) ignored", (gates / "tests.log").read_text())
    assert sum(int(row[0]) for row in totals) == receipt["test_totals_including_doctests"]["passed"]
    results = HERE / "results"
    before = results / "source-manifest-before.txt"
    after = results / "source-manifest-after.txt"
    assert before.read_bytes() == after.read_bytes()
    provenance = (results / "source-provenance.txt").read_text()
    assert f"source_manifest_before_sha256={digest(before)}" in provenance
    checked_sources = 0
    for expected, name in re.findall(r"^([a-f0-9]{64})  (.+)$", provenance, re.MULTILINE):
        # Recorded paths identify the measured checkout; map its crate suffix
        # to this checkout so the verification can be moved with the repository.
        suffix = name[name.index("crates/"):]
        assert digest(ROOT / suffix) == expected, suffix
        checked_sources += 1
    assert checked_sources >= 8
    count = 0
    rows = []
    for fixture in ["small", "opaque"]:
        for operation in ["read", "clone", "noop", "change"]:
            samples = []
            for process in range(1, 4):
                path = results / f"{fixture}-{operation}-p{process}.json"
                data = json.loads(path.read_text())
                assert data["fixture"] == fixture and data["operation"] == operation
                assert data["sample_count"] == len(data["samples"]) == 30
                for key in ["allocator_instrumented", "fixture_shape_ok", "semantic_ok_all", "opaque_preserved_all", "changed_ok_all", "inverse_ok_all"]:
                    assert data[key], (path, key)
                for sample in data["samples"]:
                    assert sample["allocated_bytes"] == sample["direct_allocated_bytes"] + sample["realloc_new_bytes"]
                    assert sample["live_after"] == sample["live_before"] + sample["direct_allocated_bytes"] + sample["realloc_new_bytes"] - sample["realloc_old_bytes"] - sample["deallocated_bytes"]
                    assert sample["peak_live_bytes"] >= max(0, sample["live_after"] - sample["live_before"])
                    assert not sample["alloc_invalid"] and sample["alloc_failed"] == 0
                    assert sample["changed_ok"] and sample["inverse_ok"]
                    if operation in ["clone", "noop"]:
                        assert sample["allocated_bytes"] == sample["alloc_calls"] == sample["realloc_calls"] == 0
                        assert sample["source_shared"]
                for percent in [50, 95, 99]:
                    assert data[f"p{percent}_ns"] == quantile([s["elapsed_ns"] for s in data["samples"]], percent)
                for prefix, field in [("allocated", "allocated_bytes"), ("peak_live", "peak_live_bytes")]:
                    for percent in [50, 95]:
                        assert data[f"{prefix}_p{percent}"] == quantile([s[field] for s in data["samples"]], percent)
                samples.extend(data["samples"])
                count += 1
            rows.append({"fixture": fixture, "operation": operation, "samples": len(samples), "p50_ns": quantile([s["elapsed_ns"] for s in samples], 50), "requested_bytes_p50": quantile([s["allocated_bytes"] for s in samples], 50)})
    result = {"passed": True, "gate_input_files": len(source_hashes), "gate_tests_including_doctests": receipt["test_totals_including_doctests"], "profile_source_manifest_sha256": digest(before), "processes": count, "samples": count * 30, "lanes": rows}
    (HERE / "root-verification.json").write_text(json.dumps(result, indent=2) + "\n")
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
