#!/usr/bin/env python3
"""Verify retained gate receipts and independently recompute profile statistics."""
import hashlib
import json
import math
from pathlib import Path

ROOT = Path(__file__).resolve().parent


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def percentile(values, fraction):
    values = sorted(values)
    position = (len(values) - 1) * fraction
    lower = math.floor(position)
    upper = math.ceil(position)
    return values[lower] + (values[upper] - values[lower]) * (position - lower)


def main():
    gates = ROOT / "gates"
    for result in json.loads((gates / "results.json").read_text()):
        assert digest(gates / (result["name"] + ".log")) == result["log_sha256"]
    assert (gates / "source-before.json").read_bytes() == (gates / "source-after.json").read_bytes()
    verification = json.loads((gates / "verification.json").read_text())
    assert verification["stable_sources"] and verification["all_required_checks_passed"]
    freeze = json.loads((gates / "freeze.json").read_text())
    gate_source = json.loads((gates / "source-after.json").read_text())
    assert all(gate_source[name] == value for name, value in freeze["selected_files"].items())
    profile = ROOT / "performance/results/final-capture-20260919"
    for name in ("source-manifest", "fixture-hashes", "fixture-member-hashes", "harness-manifest", "harness-binary"):
        assert (profile / (name + "-before.sha256")).read_bytes() == (profile / (name + "-after.sha256")).read_bytes()
    manifest = dict(line.split("  ", 1)[::-1] for line in (profile / "source-manifest-before.sha256").read_text().splitlines())
    assert all(manifest[name] == value for name, value in freeze["selected_files"].items())
    receipt = json.loads((profile / "receipt-index.json").read_text())
    for filename, key in (("raw-receipts.jsonl", "raw_sha256"), ("stats.json", "stats_sha256"), ("report.md", "report_sha256")):
        assert digest(profile / filename) == receipt[key]
    run = json.loads((profile / "run-manifest.json").read_text())
    assert digest(gates / "freeze.json") == run["freeze_receipt_sha256"]
    records = [json.loads(line) for line in (profile / "raw-receipts.jsonl").read_text().splitlines() if line.strip()]
    samples = [row for row in records if row["record"] == "sample"]
    assert len(samples) == 360
    checks = [row for row in records if row["record"] == "correctness"]
    assert len(checks) == 3
    for row in checks:
        assert all(row[key] is True for key in ("source_noop_exact", "eager_noop_exact", "source_changed_reopen", "eager_changed_reopen", "inverse_canonical_exact"))
    stats = json.loads((profile / "stats.json").read_text())
    assert len(stats["groups"]) == 24
    for group in stats["groups"]:
        rows = [row for row in samples if row["fixture"] == group["fixture"] and row["lane"] == group["lane"]]
        assert len(rows) == group["sample_count"] == 15
        for metric, actual in group["metrics"].items():
            values = [row[metric] for row in rows]
            expected = {"min": min(values), "max": max(values), **{name: percentile(values, fraction) for name, fraction in (("p50", .5), ("p95", .95), ("p99", .99))}}
            assert all(math.isclose(actual[key], value, rel_tol=1e-12, abs_tol=1e-8) for key, value in expected.items()), (group["fixture"], group["lane"], metric)
    print("Verified gate log hashes, frozen source identity, 360 samples, correctness records, and all reported statistics.")


if __name__ == "__main__":
    main()
