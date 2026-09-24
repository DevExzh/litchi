#!/usr/bin/env python3
"""Independently verify final source identity and retained gate receipts."""
import hashlib
import json
from collections import defaultdict
from pathlib import Path
import re
import statistics
import subprocess

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def verify_gates():
    gates = HERE / "gates"
    before = json.loads((gates / "source-before.json").read_text())
    after = json.loads((gates / "source-after.json").read_text())
    assert before == after, "source changed during gates"
    freeze = json.loads((gates / "freeze.json").read_text())
    for relative, expected in freeze["selected_files"].items():
        assert before[relative] == expected, relative
        if relative == "Cargo.lock":
            assert digest(gates / "Cargo.lock") == expected
        else:
            assert digest(REPO / relative) == expected, relative
    commands = json.loads((gates / "results.json").read_text())
    assert {row["name"] for row in commands} == {
        "ods-tests", "clippy", "rustdoc", "format", "batch-format", "boundaries", "diff-check"
    }
    assert len(commands) == 7
    for row in commands:
        assert row["exit_code"] == 0, row["name"]
        assert digest(gates / (row["name"] + ".log")) == row["log_sha256"]
    summaries = re.findall(
        r"test result: ok\. (\d+) passed; (\d+) failed; (\d+) ignored",
        (gates / "ods-tests.log").read_text(),
    )
    assert summaries
    counts = [sum(int(row[index]) for row in summaries) for index in range(3)]
    assert counts[1] == 0
    verification = json.loads((gates / "verification.json").read_text())
    assert verification == {"stable_sources": True, "all_required_checks_passed": True}
    return dict(zip(("passed", "failed", "ignored"), counts, strict=True))


def verify_native():
    rows = json.loads((HERE / "native/cached-results.json").read_text())
    provenance = json.loads((HERE / "native/provenance.json").read_text())
    assert len(rows) == 115
    assert len({row["function"] for row in rows}) == 24
    assert len(provenance["inputs"]) == 24
    assert len(provenance["excluded_profile_variances"]) == 1
    assert all(row["source"] in provenance["inputs"] for row in rows)
    return {"cached_observations": len(rows), "functions": 24, "profile_variances": 1}


def verify_performance():
    captures = {}
    results = HERE / "performance/results"
    retained = json.loads((results / "retained-files.json").read_text())
    assert set(retained) == {str(path.relative_to(results)) for path in results.rglob("*")
                             if path.is_file() and path.name != "retained-files.json"}
    for path, expected in retained.items():
        assert digest(results / path) == expected, path
    profile = json.loads((results / "profile-inputs-before.json").read_text())
    assert profile == json.loads((results / "profile-inputs-after.json").read_text())
    for path, expected in profile.items():
        assert digest(HERE / "performance" / path) == expected, path
    freeze = json.loads((HERE / "gates/freeze.json").read_text())
    gate_sources = json.loads((HERE / "gates/source-before.json").read_text())
    manifests = {}
    environments = {}
    for label in ("baseline-069c65870", "candidate-final"):
        directory = HERE / "performance/results" / label
        rows = [json.loads(line) for line in (directory / "measurements.jsonl").read_text().splitlines()]
        manifest = json.loads((directory / "source-manifest.json").read_text())
        before, after = manifest["before"], manifest["after"]
        assert before == after, label
        assert before["git_head"] == freeze["base_commit"]
        assert before["profile_input_sha256"] == profile
        assert before["workspace_lock_sha256"] == freeze["selected_files"]["Cargo.lock"]
        assert json.loads((directory / "target-cleanup.json").read_text())["removed"]
        if label == "candidate-final":
            for path, expected in freeze["selected_files"].items():
                assert before["source_sha256"][path] == expected, path
            for path, expected in before["workspace_source_sha256"].items():
                if path in gate_sources:
                    assert gate_sources[path] == expected, path
                else:
                    # The gate manifest excludes examples and fuzz targets.
                    assert path.split("/")[2] in ("examples", "fuzz"), path
                    content = subprocess.check_output(
                        ["git", "show", freeze["base_commit"] + ":" + path], cwd=REPO)
                    assert hashlib.sha256(content).hexdigest() == expected, path
        manifests[label] = manifest
        environments[label] = json.loads((directory / "environment.json").read_text())
        groups = defaultdict(list)
        for row in rows:
            assert row["capture"] == label and row["supported"] is True
            assert row["binary_sha256"] == manifest["binary_sha256"]
            assert row["source_git_head"] == freeze["base_commit"]
            assert row["iterations"] == 1 and row["warmups"] == 3
            assert len(row["samples"]) == 1 and row["repeat"] > 0
            sample = row["samples"][0]
            assert sample["elapsed_ns"] > 0
            assert sample["live_before"] == sample["live_after"]
            assert sample["requested_bytes"] == sample["released_bytes"]
            original = json.loads((directory / row["raw_stdout"]).read_text())
            for key, value in original.items():
                if key != "rss_kib":
                    assert row[key] == value, (label, row["case"], key)
            time_log = (directory / row["raw_time"]).read_text()
            assert "Exit status: 0" in time_log
            rss = int(re.search(r"Maximum resident set size \(kbytes\):\s*(\d+)", time_log)[1])
            assert row["rss_kib"] == rss
            groups[(row["case"], row["phase"])].append(row)
        medians = {}
        for key, samples in groups.items():
            assert len(samples) == 15
            assert {row["sample_index"] for row in samples} == set(range(1, 16))
            for stable in ("repeat", "input_bytes", "operation", "shape", "rows", "columns", "elements"):
                assert len({row[stable] for row in samples}) == 1, (key, stable)
            assert len({row["samples"][0]["checksum"] for row in samples}) == 1, key
            medians[key] = {
                metric: statistics.median(row["samples"][0][metric] for row in samples)
                for metric in ("elapsed_ns", "alloc_calls", "requested_bytes", "peak_live_delta", "work", "memory_retained")
            }
            medians[key]["rss_kib"] = statistics.median(row["rss_kib"] for row in samples)
        captures[label] = {"rows": rows, "groups": groups, "medians": medians}
    baseline, candidate = captures["baseline-069c65870"], captures["candidate-final"]
    left_manifest, right_manifest = (manifests[label]["before"] for label in ("baseline-069c65870", "candidate-final"))
    for field in ("harness_sha256", "fixture_sha256", "workspace_lock_sha256"):
        assert left_manifest[field] == right_manifest[field], field
    for field in ("rustc_verbose", "cargo", "libc", "rustflags"):
        assert environments["baseline-069c65870"][field] == environments["candidate-final"][field], field
    before_sources, after_sources = left_manifest["workspace_source_sha256"], right_manifest["workspace_source_sha256"]
    changed = {path for path in set(before_sources) | set(after_sources) if before_sources.get(path) != after_sources.get(path)}
    assert changed == {path for path in freeze["selected_files"] if path.endswith(".rs")}, changed
    controls = {key for key, rows in candidate["groups"].items() if rows[0]["operation"] in ("ARITHMETIC", "ROUND")}
    assert set(baseline["groups"]) == controls and len(controls) == 20
    report = json.loads((results / "performance-report.json").read_text())
    assert report["baseline_groups"] == 20
    assert report["candidate_groups"] == len(candidate["groups"])
    reported = {"baseline-069c65870": {}, "candidate-final": {}}
    for pair in report["controls"]:
        for side, label in (("baseline", "baseline-069c65870"), ("candidate", "candidate-final")):
            row = pair[side]
            key = (row["case"], row["phase"])
            assert key not in reported[label]
            reported[label][key] = row
    for row in report["candidate_only"]:
        key = (row["case"], row["phase"])
        assert key not in reported["candidate-final"]
        reported["candidate-final"][key] = row
    metrics = {"elapsed_ns": "time_ns_batch_p50", "alloc_calls": "allocator_calls",
               "requested_bytes": "requested_bytes", "peak_live_delta": "peak_live_bytes",
               "memory_retained": "result_live_memory", "work": "work", "rss_kib": "rss_kib"}
    for label, capture in captures.items():
        assert set(reported[label]) == set(capture["groups"])
        for key, medians in capture["medians"].items():
            row = reported[label][key]
            repeat = capture["groups"][key][0]["repeat"]
            assert row["samples"] == 15 and row["repeat"] == repeat
            for metric, field in metrics.items():
                assert row[field] == medians[metric], (label, key, metric)
            assert row["time_ns_per_repeat"] == medians["elapsed_ns"] // repeat
    comparisons = {}
    for key in sorted(controls):
        left, right = baseline["groups"][key][0], candidate["groups"][key][0]
        for field in ("repeat", "input_bytes", "operation", "shape", "rows", "columns", "elements"):
            assert left[field] == right[field], (key, field)
        assert left["samples"][0]["checksum"] == right["samples"][0]["checksum"], key
        before, after = baseline["medians"][key], candidate["medians"][key]
        comparisons["/".join(key)] = {
            metric: {"baseline": before[metric], "candidate": after[metric],
                     "ratio": after[metric] / before[metric] if before[metric] else None}
            for metric in before
        }
    return {"baseline_samples": len(baseline["rows"]), "candidate_samples": len(candidate["rows"]),
            "candidate_groups": len(candidate["groups"]), "control_comparisons": comparisons}


if __name__ == "__main__":
    print(json.dumps({"gates": verify_gates(), "native": verify_native(),
                      "performance": verify_performance()}, indent=2))
