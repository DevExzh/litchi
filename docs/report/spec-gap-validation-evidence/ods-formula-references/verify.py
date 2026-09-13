#!/usr/bin/env python3
"""Verify retained ODS reference test and benchmark evidence against source."""
import csv
import hashlib
import json
from pathlib import Path
import re
import subprocess

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

receipt = json.loads((HERE / "gates/results.json").read_text())
assert receipt["sources_unchanged"]
assert receipt["source_before"] == receipt["source_after"]
for name, expected in receipt["source_after"].items():
    assert sha(ROOT / name) == expected, name
assert len(receipt["commands"]) == 5
assert all(x["status"] == 0 for x in receipt["commands"])
results = re.findall(r"test result: ok\. (\d+) passed; (\d+) failed; (\d+) ignored", (HERE / "gates/test.log").read_text())
assert len(results) == 47
assert sum(int(x[0]) for x in results) == 810
assert all(x[1:] == ("0", "0") for x in results)
replay = json.loads((HERE / "gates/root-patch-replay.json").read_text())
assert replay["status"] == 0 and replay["base"] == receipt["head"]
assert len(replay["replayed_sha256"]) == 5
assert all(receipt["source_after"][name] == value for name, value in replay["replayed_sha256"].items())
probe = json.loads((HERE / "gates/iri-probe.json").read_text())
assert probe["build_exit_code"] == probe["run_exit_code"] == probe["allocation_calls"] == 0
for name, expected in probe["source_sha256"].items():
    assert sha(ROOT / name) == expected
assert (HERE / "gates/iri-probe.log").read_text().strip() == "cases=37 observations=37016 long_input_bytes=65536 allocations=0"

baseline = HERE / "performance/baseline"
for line in (baseline / "source-sha256.txt").read_text().splitlines():
    expected, name = line.split("  ", 1)
    content = subprocess.check_output(["git", "show", receipt["head"] + ":" + name], cwd=ROOT)
    assert hashlib.sha256(content).hexdigest() == expected, name

candidate = HERE / "performance/candidate"
assert json.loads((candidate / "source-sha256.json").read_text()) == receipt["source_after"]
assert (candidate / "patch-replay.status").read_text().strip() == "0"
checks = json.loads((candidate / "isolated-checks.json").read_text())
assert checks["all_passed"]
assert sum(x["passed"] for x in checks["checks"]) == 64
binaries = json.loads((HERE / "gates/root-binary-verification.json").read_text())["binaries"]
for kind in ["baseline", "candidate"]:
    expected, name = (HERE / "performance" / kind / "binary-sha256.txt").read_text().strip().split("  ", 1)
    assert binaries[kind] == {"path": name, "sha256": expected}
for directory in [baseline, candidate]:
    for line in (directory / "harness-sha256.txt").read_text().splitlines():
        expected, name = line.split("  ", 1)
        assert sha(ROOT / name) == expected, name

lane_count = 0
case_sets = []
for group_dir in ["baseline", "candidate", "coverage"]:
    directory = HERE / "performance" / group_dir
    rows = list(csv.DictReader((directory / "raw.csv").open()))
    assert len(rows) == (16 if group_dir == "coverage" else 12)
    case_sets.append([r["case"] for r in rows])
    for row in rows:
        stem = directory / ("parse-" + row["case"])
        lines = stem.with_suffix(".stdout").read_text().splitlines()
        parsed = {}
        for prefix in ["config ", "result "]:
            line = next(x for x in lines if x.startswith(prefix))
            parsed.update(dict(x.split("=", 1) for x in line.split()[1:]))
        for key, value in parsed.items():
            if key in row:
                assert row[key] == value, (group_dir, row["case"], key)
        time_text = stem.with_suffix(".time").read_text()
        rss = re.search(r"Maximum resident set size \(kbytes\):\s*(\d+)", time_text).group(1)
        assert row["max_rss_kib"] == rss
        assert row["status"] == stem.with_suffix(".status").read_text().strip() == "0"
        assert stem.with_suffix(".stderr").read_text() == ""
        expected = int(row["repeat"]) if row["expected_success"] == "true" else 0
        assert int(row["successes_p50"]) == int(row["successes_max"]) == expected
        assert row["iterations"] == "15" and row["warmups"] == "3"
        lane_count += 1
assert case_sets[0] == case_sets[1]
manifest = HERE / "artifacts.sha256"
manifest_lines = manifest.read_text().splitlines()
assert len(manifest_lines) >= 180
for line in manifest_lines:
    expected, name = line.split("  ", 1)
    assert sha(HERE / name) == expected, name
print(json.dumps({"verified": True, "tests": 810, "targets": 47, "performance_lanes": lane_count, "iri_observations": 37016, "iri_allocation_calls": 0}))
