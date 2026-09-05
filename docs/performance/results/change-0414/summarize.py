#!/usr/bin/env python3
"""Recompute descriptive ZIP32 guards; this does not authorize a speedup claim."""
import hashlib
import atexit
import json
import math
from pathlib import Path
import re
import subprocess
import sys
import tempfile

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
LEGS = ("A1", "B1", "B2", "A2")
identity = json.loads((ROOT / "identities.json").read_text())
for relative, digest in identity["candidate_sources"].items():
    source = subprocess.run(
        ["git", "-C", str(REPO), "show", identity["candidate_revision"] + ":" + relative],
        stdout=subprocess.PIPE, stderr=subprocess.DEVNULL,
    )
    raw = source.stdout if source.returncode == 0 else (REPO / relative).read_bytes()
    assert hashlib.sha256(raw).hexdigest() == digest
scratch = tempfile.TemporaryDirectory(prefix="litchi-goal-0414-replay-")
atexit.register(scratch.cleanup)

def materialize(relative):
    """Verify and materialize a retained lossless original for the existing validator."""
    index = json.loads((ROOT / "raw-artifacts.json").read_text())
    raw = subprocess.check_output(["zstd", "-q", "-d", "-c", str(ROOT / (relative + ".zst"))])
    assert len(raw) == index[relative]["bytes"]
    assert hashlib.sha256(raw).hexdigest() == index[relative]["sha256"]
    target = Path(scratch.name) / relative
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_bytes(raw)
    return target

protocol_bytes = (ROOT / "protocol.json").read_bytes()
assert hashlib.sha256(protocol_bytes).hexdigest() == identity["protocol_sha256"]
protocol = json.loads(protocol_bytes)
assert sys.argv[1:] in ([], ["--followup"])
capture_dir = "guards"
if sys.argv[1:]:
    protocol_bytes = (ROOT / "guard-recheck-protocol.json").read_bytes()
    assert hashlib.sha256(protocol_bytes).hexdigest() == identity["guard_recheck_protocol_sha256"]
    protocol = json.loads(protocol_bytes)
    capture_dir = "guard-recheck"
reports, rss, rows = {}, {}, {}
expected_keys = {
    (case, f"{shape}-{payload}")
    for case in protocol["guard_cases"]
    for shape in protocol["guard_shapes"]
    for payload in protocol["guard_payloads"]
}
for leg in LEGS:
    report_path = materialize(f"{capture_dir}/{leg}.json")
    catalog_path = materialize(f"{capture_dir}/{leg}.corpus.json")
    subprocess.run([
        sys.executable, str(REPO / "tools/validate_perf_corpus_binding.py"),
        "--report", str(report_path), "--catalog",
        str(catalog_path),
    ], check=True, stdout=subprocess.DEVNULL)
    report = json.loads(report_path.read_text())
    role = "control" if leg.startswith("A") else "candidate"
    assert report["environment"]["git_revision"] == identity[f"{role}_revision"]
    assert report["environment"]["git_worktree_dirty"] is False
    assert report["environment"]["cpu_affinity"] == str(protocol["cpu"])
    assert report["configuration"]["samples_per_case"] == protocol["samples_per_row"]
    assert report["configuration"]["warmup_iterations_per_case"] == protocol["warmups_per_row"]
    keyed = {(r["case"], r["corpus"]["name"]): r for r in report["results"]}
    assert set(keyed) == expected_keys and len(keyed) == len(report["results"])
    for key, row in keyed.items():
        stats = row["elapsed_ns"]
        samples = stats["samples"]
        n = protocol["samples_per_row"]
        assert len(samples) == n and sorted(samples) == samples
        assert all(type(value) is int and value >= 0 for value in samples)
        assert sorted(stats["sample_order"]) == list(range(n))
        assert stats["p50"] == (samples[(n - 1) // 2] + samples[n // 2]) // 2
        for percent in (95, 99):
            assert stats[f"p{percent}"] == samples[math.ceil(n * percent / 100) - 1]
        if leg != "A1":
            assert row["corpus"] == rows["A1"][key]["corpus"]
            assert row["sink"] == rows["A1"][key]["sink"]
    reports[leg], rows[leg] = report, keyed
    time_text = (ROOT / capture_dir / f"{leg}.time.txt").read_text()
    rss[leg] = int(re.search(r"Maximum resident set size \(kbytes\):\s*(\d+)", time_text)[1])
    assert "Exit status: 0" in time_text
for first, second in (("A1", "A2"), ("B1", "B2")):
    assert reports[first]["binary_identity"]["binary_sha256"] == reports[second]["binary_identity"]["binary_sha256"]

def changes(values):
    return [(values[b] / values[a] - 1) * 100 for a, b in (("A1", "B1"), ("A2", "B2"))]

result = {"performance_claim": "none", "rows": [], "peak_rss_kib": rss,
          "peak_rss_change_percent": changes(rss), "review_triggers": []}
for key in sorted(expected_keys):
    row = {"case": key[0], "corpus": key[1], "metrics": {}}
    for metric in ("p50", "p95", "p99"):
        values = {leg: rows[leg][key]["elapsed_ns"][metric] for leg in LEGS}
        delta = changes(values)
        row["metrics"][metric] = {"ns": values, "change_percent": delta}
        for pair, change in enumerate(delta, 1):
            if change > 5:
                result["review_triggers"].append([*key, metric, pair, change])
    result["rows"].append(row)
for pair, change in enumerate(result["peak_rss_change_percent"], 1):
    if change > 5:
        result["review_triggers"].append(["process_peak_rss", pair, change])
print(json.dumps(result, indent=2, allow_nan=False))
