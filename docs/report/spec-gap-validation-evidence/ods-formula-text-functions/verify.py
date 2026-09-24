#!/usr/bin/env python3
"""Independently verify text source custody, raw measurements, and gate evidence."""
import hashlib
import json
from collections import defaultdict
from pathlib import Path
import random
import re
import statistics
import subprocess

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
FUNCTIONS = set("ASC CHAR CLEAN CODE CONCATENATE DOLLAR EXACT FIND FIXED JIS LEFT LEN LOWER MID PROPER REPLACE REPT RIGHT SEARCH SUBSTITUTE T TEXT TRIM UNICHAR UNICODE UPPER".split())


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def verify_gates():
    run = subprocess.run(["python3", str(HERE / "gates/verify.py")],
                         check=True, capture_output=True, text=True)
    return json.loads(run.stdout)


def verify_oracles():
    run = subprocess.run(["python3", str(HERE / "text_oracle.py"), "--check"],
                         check=True, capture_output=True, text=True)
    document = json.loads((HERE / "text-goldens.json").read_text())
    rows = document["observations"]
    assert document["contract_sha256"] == digest(HERE / "contract.md")
    assert {row["function"] for row in rows} == FUNCTIONS
    receipt = json.loads(run.stdout)
    assert receipt == {"functions": 26, "observations": len(rows), "verified": True}
    return receipt


def verify_native():
    rows = json.loads((HERE / "native/cached-results.json").read_text())
    provenance = json.loads((HERE / "native/provenance.json").read_text())
    assert len(rows) == provenance["selected_observations"] == 91
    assert {row["function"] for row in rows} == FUNCTIONS
    assert len(provenance["inputs"]) == 26
    assert all(row["source"] in provenance["inputs"] for row in rows)
    assert sum(row["valid_modes"] == ["matrix"] for row in rows) == 7
    receipt = json.loads((HERE / "native-reproduction.json").read_text())
    assert receipt["files"] == 26 and receipt["temporary_tree_cleaned"]
    assert receipt["selected_observations"] == 91 and not receipt["converter"]
    assert not receipt["recalculation"] and receipt["commit"] == provenance["commit"]
    custody = json.loads((HERE / "native-reproduction-custody.json").read_text())
    assert custody["exit_code"] == 0 and not custody["stderr"]
    assert custody["before"] == custody["after"]
    assert json.loads(custody["stdout"]) == receipt
    for path, expected in custody["before"].items():
        assert digest(HERE / path) == expected, path
    return {"observations": 91, "functions": 26, "matrix_only": 7}


def verify_unicode():
    provenance = json.loads((HERE / "unicode-data/provenance.json").read_text())
    receipt = json.loads((HERE / "unicode-reproduction.json").read_text())
    assert receipt["verified"] and receipt["temporary_tree_cleaned"]
    assert receipt["unicode_version"] == provenance["unicode_version"] == "17.0.0"
    assert digest(REPO / provenance["generated"]["path"]) == receipt["output_sha256"] == provenance["generated"]["sha256"]
    assert digest(REPO / provenance["generator"]["path"]) == provenance["generator"]["sha256"]
    assert digest(HERE / "unicode-data" / provenance["license"]["path"]) == provenance["license"]["sha256"]
    return receipt


def verify_fraction():
    receipt = json.loads((HERE / "fraction-reproduction.json").read_text())
    assert receipt["verified"] and receipt["temporary_tree_cleaned"]
    assert receipt["cases"] == 5474 and receipt["seed"] == 20260920
    source = REPO / "crates/litchi-ods/src/codec/formula/evaluation/text/fraction.rs"
    assert digest(source) == receipt["source_sha256"]
    assert digest(HERE / "verify_fraction.py") == receipt["verifier_sha256"]
    run = subprocess.run(["python3", str(HERE / "verify_fraction.py"), "--check"],
                         check=True, capture_output=True, text=True)
    assert json.loads(run.stdout)["output_sha256"] == receipt["output_sha256"]
    return receipt


def verify_superseded_gates():
    directory = HERE / "diagnostics/pre-grammar-fix"
    retained = json.loads((directory / "retained-files.json").read_text())
    assert set(retained) == {
        str(path.relative_to(directory)) for path in directory.rglob("*")
        if path.is_file() and path.name != "retained-files.json"
    }
    for relative, expected in retained.items():
        assert digest(directory / relative) == expected, relative
    gates = directory / "gates"
    freeze = json.loads((gates / "freeze.json").read_text())
    before = json.loads((gates / "source-before.json").read_text())
    assert before == json.loads((gates / "source-after.json").read_text())
    for relative, expected in freeze["selected_files"].items():
        assert digest(directory / "source" / relative) == before[relative] == expected
    results = json.loads((gates / "results.json").read_text())
    assert len(results) == 7 and all(row["exit_code"] == 0 for row in results)
    for row in results:
        assert digest(gates / (row["name"] + ".log")) == row["log_sha256"]
    return {"verified": True, "selected_sources": len(freeze["selected_files"]),
            "disposition": "superseded by uncovered formatter grammar findings"}


def verify_performance():
    captures = {}
    results = HERE / "performance/results"
    retained = json.loads((results / "retained-files.json").read_text())
    assert set(retained) == {str(path.relative_to(results)) for path in results.rglob("*")
                             if path.is_file() and path != results / "retained-files.json"}
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
    for label in ("baseline-8f09231e36982248eface4d599432143a67f6e49", "candidate-final"):
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
                assert before["source_sha256"].get(path) == expected, path
            for path, expected in before["source_sha256"].items():
                current = HERE / "gates/Cargo.lock" if path == "Cargo.lock" else REPO / path
                assert (digest(current) if current.is_file() else None) == expected, path
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
            assert row["output_bytes_p50"] == sample["output_bytes"]
            assert row["bytes_per_repeat_p50"] == row["input_bytes"] + sample["output_bytes"] // row["repeat"]
            assert row["validation_scope"].startswith("one untimed contract fixture oracle")
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
            case = row["case"]
            cells = row["rows"] * row["columns"]
            expected = expected_reads(case, cells)
            if expected is not None:
                if case.startswith("cancellation-text-"):
                    # One context is shared by the four internal repeats.
                    # The first call cancels on its first read; subsequent
                    # calls refuse before reading. This is an exact batch
                    # count, not a relaxed per-repeat read interval.
                    assert row["repeat"] == 4 and row["expected"] == "failure:cancelled"
                    assert sample["reference_reads"] == 1, case
                    assert row["reference_reads_per_repeat"] == 0, case
                else:
                    assert sample["reference_reads"] == expected * row["repeat"], (case, sample["reference_reads"], expected)
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
                for metric in ("elapsed_ns", "alloc_calls", "requested_bytes", "released_bytes", "peak_live_delta", "work", "memory_retained", "reference_reads", "output_bytes")
            }
            medians[key]["rss_kib"] = statistics.median(row["rss_kib"] for row in samples)
        captures[label] = {"rows": rows, "groups": groups, "medians": medians}
    baseline, candidate = captures["baseline-8f09231e36982248eface4d599432143a67f6e49"], captures["candidate-final"]
    left_manifest, right_manifest = (manifests[label]["before"] for label in ("baseline-8f09231e36982248eface4d599432143a67f6e49", "candidate-final"))
    for field in ("harness_sha256", "workspace_lock_sha256"):
        assert left_manifest[field] == right_manifest[field], field
    for field in ("rustc_verbose", "cargo", "libc", "rustflags"):
        assert environments["baseline-8f09231e36982248eface4d599432143a67f6e49"][field] == environments["candidate-final"][field], field
    before_sources, after_sources = left_manifest["workspace_source_sha256"], right_manifest["workspace_source_sha256"]
    changed = {path for path in set(before_sources) | set(after_sources) if before_sources.get(path) != after_sources.get(path)}
    assert changed == {path for path in freeze["selected_files"] if path.endswith(".rs")}, changed
    control_names, cases = expected_cases()
    controls = {key for key in candidate["groups"] if key[0] in control_names}
    expected_controls = {(name, phase) for name in control_names
                         for phase in ("evaluate", "parse-evaluate")}
    assert set(baseline["groups"]) == controls == expected_controls
    assert set(candidate["groups"]) == {(name, phase) for name in cases
                                        for phase in ("evaluate", "parse-evaluate")}
    summary = json.loads((results / "capture-summary.json").read_text())
    preflight = summary["candidate_preflight_gate"]
    assert preflight["status"] == "ok" and preflight["cases"] == len(cases)
    assert preflight["profile_input_sha256"] == profile
    assert preflight["binary_sha256"] == manifests["candidate-final"]["binary_sha256"]
    assert set(preflight["preflight"]["cases"]) == cases
    preflight_dir = results / "preflight-before-timing"
    assert json.loads((preflight_dir / "preflight.json").read_text()) == preflight["preflight"]
    assert json.loads((preflight_dir / "target-cleanup.json").read_text())["removed"]
    output = (preflight_dir / "preflight.stdout.log").read_text()
    observations = re.findall(r"^preflight case=(\S+) reference_reads=(\d+)$", output, re.MULTILINE)
    assert len(observations) == len(cases) and {case for case, _ in observations} == cases
    assert output.count("preflight-ok cases=1 phase=evaluate") == len(cases)
    for case, reads in observations:
        row = candidate["groups"][(case, "evaluate")][0]
        expected = expected_reads(case, row["rows"] * row["columns"])
        if expected is not None:
            assert int(reads) == expected, (case, reads, expected)
    report = json.loads((results / "performance-report.json").read_text())
    assert report["baseline_groups"] == 2 * len(control_names)
    assert report["candidate_groups"] == len(candidate["groups"]) == 2 * len(cases)
    reported = {"baseline-8f09231e36982248eface4d599432143a67f6e49": {}, "candidate-final": {}}
    for pair in report["controls"]:
        for side, label in (("baseline", "baseline-8f09231e36982248eface4d599432143a67f6e49"), ("candidate", "candidate-final")):
            row = pair[side]
            key = (row["case"], row["phase"])
            assert key not in reported[label]
            reported[label][key] = row
    for row in report["candidate_only"]:
        key = (row["case"], row["phase"])
        assert key not in reported["candidate-final"]
        reported["candidate-final"][key] = row
    metrics = {"alloc_calls": "allocator_calls",
               "requested_bytes": "requested_bytes", "peak_live_delta": "peak_live_bytes",
               "memory_retained": "result_live_budget", "released_bytes": "released_bytes",
               "rss_kib": "rss_kib", "output_bytes": "output_bytes_p50"}
    for label, capture in captures.items():
        assert set(reported[label]) == set(capture["groups"])
        for key, medians in capture["medians"].items():
            row = reported[label][key]
            repeat = capture["groups"][key][0]["repeat"]
            assert row["samples"] == 15 and row["repeat"] == repeat
            for metric, field in metrics.items():
                assert row[field] == medians[metric], (label, key, metric)
            assert row["time_ns_per_repeat"] == medians["elapsed_ns"] // repeat
            assert row["work_per_repeat"] == medians["work"] // repeat
            assert row["reference_reads"] == medians["reference_reads"] // repeat
    comparisons = {}
    for key in sorted(controls):
        left, right = baseline["groups"][key][0], candidate["groups"][key][0]
        for field in ("repeat", "input_bytes", "operation", "shape", "rows", "columns", "elements"):
            assert left[field] == right[field], (key, field)
        assert left["samples"][0]["checksum"] == right["samples"][0]["checksum"], key
        before, after = baseline["medians"][key], candidate["medians"][key]
        for metric in before.keys() - {"elapsed_ns", "rss_kib"}:
            assert before[metric] == after[metric], (key, metric)
        comparisons["/".join(key)] = {
            metric: {"baseline": before[metric], "candidate": after[metric],
                     "ratio": after[metric] / before[metric] if before[metric] else None}
            for metric in before
        }
    flags = {}
    for key, comparison in comparisons.items():
        for metric in ("elapsed_ns", "rss_kib"):
            if comparison[metric]["ratio"] <= 1.05:
                continue
            group = tuple(key.split("/"))
            values = [[row["rss_kib"] if metric == "rss_kib" else row["samples"][0][metric]
                       for row in sorted(capture["groups"][group], key=lambda r: r["sample_index"])]
                      for capture in (baseline, candidate)]
            rng = random.Random(20260920)
            bootstrap = sorted(statistics.median(rng.choices(values[1], k=15)) /
                               statistics.median(rng.choices(values[0], k=15)) - 1
                               for _ in range(10000))
            flags[key + "/" + metric] = {
                "ratio": comparison[metric]["ratio"],
                "baseline_range": [min(values[0]), max(values[0])],
                "candidate_range": [min(values[1]), max(values[1])],
                "bootstrap_percentile_95_delta": [bootstrap[249], bootstrap[9749]],
            }
    return {"baseline_samples": len(baseline["rows"]), "candidate_samples": len(candidate["rows"]),
            "candidate_groups": len(candidate["groups"]), "control_comparisons": comparisons,
            "threshold_flags": flags,
            "bootstrap": {"seed": 20260920, "resamples": 10000,
                          "method": "Independent 15-sample median resampling; candidate then baseline; indices 249 and 9749",
                          "limitation": "Single-capture uncertainty, not causal or cross-platform evidence"}}


def expected_cases():
    controls = {
        "scalar-control-arithmetic", "scalar-control-sin", "scalar-control-imsum",
        "scalar-control-average", "scalar-control-counta", "scalar-control-var",
        "scalar-control-stdev", "database-control-dsum", "database-control-dvar",
        "database-control-dstdev", "array-control-4x4-arithmetic",
        "array-control-4x4-sin", "array-control-16x16-arithmetic",
        "array-control-16x16-sin", "reference-array-16x4-arithmetic",
        "scalar-aggregate-sum", "literal-aggregate-4x1-sum",
        "reference-aggregate-64x4-sum", "reference-conditional-256x4-sumifs",
        "reference-control-average", "reference-control-counta",
        "representative-median", "representative-rank", "representative-percentrank",
    }
    controls |= {"concat-borrowed-literals", "concat-owned-left", "concat-owned-right", "concat-growth-chain"}
    cases = controls | {f"{lane}-text-{function.lower()}" for function in FUNCTIONS
                        for lane in ("tiny", "large-unicode", "reference-64", "refusal")}
    cases |= {f"matrix-broadcast-text-{function}" for function in
              ("concatenate", "exact", "find", "left", "len", "lower", "mid", "proper", "replace", "right", "search", "substitute", "trim")}
    cases |= {f"{lane}-text-{function}" for lane in ("cancellation", "resource")
              for function in ("concatenate", "len", "substitute", "search")}
    cases |= {f"search-worstcase-{function}" for function in ("find", "search", "substitute", "exact")}
    cases |= {"rept-growth", "asc-jis-expansion-asc", "asc-jis-expansion-jis",
              "text-format-fraction-six"}
    assert len(controls) == 28 and len(cases) == 161
    return controls, cases


def expected_reads(case, cells):
    if case.startswith("cancellation-text-"):
        return 1
    if case.startswith("reference-64-text-"):
        return 128 if case.endswith("-exact") else 64
    if case.startswith("matrix-broadcast-text-"):
        function = case.removeprefix("matrix-broadcast-text-")
        return 128 if function in {"concatenate", "exact", "left", "right", "mid", "replace"} else 64
    if case == "reference-conditional-256x4-sumifs":
        return 1792
    if case.startswith("database-"):
        return 7
    if case.startswith("reference-"):
        return cells
    return 0


if __name__ == "__main__":
    print(json.dumps({"gates": verify_gates(), "native": verify_native(),
                      "unicode": verify_unicode(), "oracles": verify_oracles(),
                      "fraction": verify_fraction(),
                      "superseded_gates": verify_superseded_gates(),
                      "performance": verify_performance()}, indent=2))
