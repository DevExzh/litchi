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
        assert before.get(relative) == expected, relative
        if expected is None:
            assert not (REPO / relative).exists(), relative
            continue
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



def verify_oracles():
    run = subprocess.run(["python3", str(HERE / "numeric_oracle.py"), "--check"],
                         check=True, capture_output=True, text=True)
    rows = json.loads((HERE / "numeric-goldens.json").read_text())["observations"]
    assert len(rows) == 576
    assert {row["function"] for row in rows} == {
        "COUNT", "COUNTA", "COUNTBLANK", "AVERAGE", "AVERAGEA", "MIN", "MAX", "MINA", "MAXA"}
    assert json.loads(run.stdout) == {"observations": 576, "verified": True}
    return {"observations": len(rows), "functions": 9, "modes_per_observation": 2}


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
    for label in ("baseline-f7fe857007b7bcb65e0512c5b5caf9135fe74a5a", "candidate-final"):
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
            assert row["validation_scope"].startswith("one untimed direct f64 fixture oracle")
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
                for metric in ("elapsed_ns", "alloc_calls", "requested_bytes", "released_bytes", "peak_live_delta", "work", "memory_retained", "reference_reads")
            }
            medians[key]["rss_kib"] = statistics.median(row["rss_kib"] for row in samples)
        captures[label] = {"rows": rows, "groups": groups, "medians": medians}
    baseline, candidate = captures["baseline-f7fe857007b7bcb65e0512c5b5caf9135fe74a5a"], captures["candidate-final"]
    left_manifest, right_manifest = (manifests[label]["before"] for label in ("baseline-f7fe857007b7bcb65e0512c5b5caf9135fe74a5a", "candidate-final"))
    for field in ("harness_sha256", "workspace_lock_sha256"):
        assert left_manifest[field] == right_manifest[field], field
    for field in ("rustc_verbose", "cargo", "libc", "rustflags"):
        assert environments["baseline-f7fe857007b7bcb65e0512c5b5caf9135fe74a5a"][field] == environments["candidate-final"][field], field
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
    report = json.loads((results / "performance-report.json").read_text())
    assert report["baseline_groups"] == 2 * len(control_names)
    assert report["candidate_groups"] == len(candidate["groups"]) == 2 * len(cases)
    reported = {"baseline-f7fe857007b7bcb65e0512c5b5caf9135fe74a5a": {}, "candidate-final": {}}
    for pair in report["controls"]:
        for side, label in (("baseline", "baseline-f7fe857007b7bcb65e0512c5b5caf9135fe74a5a"), ("candidate", "candidate-final")):
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
               "rss_kib": "rss_kib"}
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
        comparisons["/".join(key)] = {
            metric: {"baseline": before[metric], "candidate": after[metric],
                     "ratio": after[metric] / before[metric] if before[metric] else None}
            for metric in before
        }
    return {"baseline_samples": len(baseline["rows"]), "candidate_samples": len(candidate["rows"]),
            "candidate_groups": len(candidate["groups"]), "control_comparisons": comparisons}



def verify_native():
    rows = json.loads((HERE / "native/cached-results.json").read_text())
    provenance = json.loads((HERE / "native/provenance.json").read_text())
    assert rows and len(rows) == provenance["selected_observations"]
    functions = {row["function"] for row in rows}
    assert functions == {"COUNT", "COUNTA", "COUNTBLANK", "AVERAGE", "AVERAGEA", "MIN", "MAX", "MINA", "MAXA"}
    assert all(row["source"] in provenance["inputs"] for row in rows)
    reproduction = json.loads((HERE / "native-reproduction.json").read_text())
    assert reproduction["exit_code"] == 0 and not reproduction["stderr"]
    assert reproduction["before"] == reproduction["after"]
    for path, expected in reproduction["before"].items():
        assert digest(HERE / path) == expected, path
    receipt = json.loads(reproduction["stdout"])
    assert receipt["files"] == len(provenance["inputs"]) and receipt["temporary_tree_cleaned"]
    assert receipt["selected_observations"] == len(rows)
    assert receipt["commit"] == provenance["commit"]
    return {"cached_observations": len(rows), "functions": len(functions),
            "exclusions": len(provenance["excluded_profile_variances"])}


def expected_cases():
    controls = {
        "scalar-control-arithmetic", "scalar-control-sin", "scalar-control-imsum",
        "database-control-dsum", "array-control-4x4-arithmetic", "array-control-4x4-sin",
        "array-control-16x16-arithmetic", "array-control-16x16-sin",
        "reference-array-16x4-arithmetic", "scalar-aggregate-sum",
        "literal-aggregate-4x1-sum", "reference-aggregate-64x4-sum",
        "reference-conditional-256x4-sumifs",
    }
    functions = {"count", "counta", "countblank", "average", "averagea", "min", "max", "mina", "maxa"}
    cases = controls | {
        f"{kind}-statistical-{function}" for function in functions
        for kind in ("scalar", "literal", "reference", "list", "3d", "empty", "error", "resource")
    }
    cases |= {f"nested-projected-statistical-64-{function}" for function in functions}
    cases |= {f"nested-projected-statistical-{rows}-{function}"
              for rows in (256, 1024) for function in ("average", "counta", "countblank")}
    assert len(controls) == 13 and len(cases) == 100
    return controls, cases


def expected_reads(case, cells):
    if case.startswith("nested-projected-statistical-"):
        return 2 * cells  # Outer IF condition once per cell, one cached inner scan.
    if case == "scalar-statistical-countblank":
        return 1
    if case.startswith("3d-statistical-"):
        return 3
    if case == "list-statistical-average" or case.startswith("resource-statistical-"):
        return 0
    if case.startswith(("list-statistical-", "reference-statistical-", "empty-statistical-", "error-statistical-")):
        return cells
    if case == "reference-conditional-256x4-sumifs":
        return 7 * cells // 4
    if case.startswith("reference-"):
        return cells
    if case.startswith(("scalar-", "literal-", "array-")):
        return 0
    return None  # DSUM is independently checked by the profile verifier.


if __name__ == "__main__":
    print(json.dumps({"gates": verify_gates(), "native": verify_native(), "oracles": verify_oracles(),
                      "performance": verify_performance()}, indent=2))
