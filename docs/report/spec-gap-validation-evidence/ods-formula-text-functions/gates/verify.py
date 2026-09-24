#!/usr/bin/env python3
"""Verify frozen source custody and the seven retained integration gates."""

import hashlib
import json
from pathlib import Path
import re

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[4]


def read(name):
    return json.loads((HERE / name).read_text())


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main():
    freeze = read("freeze.json")
    baseline = json.loads((HERE.parent / "baseline.json").read_text())["commit"]
    assert freeze["base_commit"] == baseline
    environment = read("environment.json")
    assert environment["head"] == baseline
    assert environment["RUSTFLAGS"] is None
    assert environment["RUSTDOCFLAGS"] == "-D warnings"
    before = read("source-before.json")
    assert before == read("source-after.json"), "sources changed during gates"
    staged = read("staged-profile-sources.json")
    for relative, expected in freeze["selected_files"].items():
        assert expected is not None, relative
        assert staged[relative] == before[relative] == expected, relative
        source = HERE / "Cargo.lock" if relative == "Cargo.lock" else ROOT / relative
        assert digest(source) == expected, relative
    rust = read("batch-files.json")
    assert len(rust) == len(set(rust))
    assert all(relative in freeze["selected_files"] for relative in rust)
    expected_commands = {
        "ods-tests": ["cargo", "test", "--locked", "--offline", "-p", "litchi-ods"],
        "clippy": ["cargo", "clippy", "--locked", "--offline", "-p", "litchi-ods", "--all-targets", "--", "-D", "warnings"],
        "rustdoc": ["cargo", "doc", "--locked", "--offline", "-p", "litchi-ods", "--no-deps"],
        "format": ["cargo", "fmt", "-p", "litchi-ods", "--", "--check"],
        "batch-format": ["rustfmt", "--edition", "2024", "--check", "--config", "skip_children=true", *rust],
        "boundaries": ["python3", "tools/check_crate_boundaries.py"],
        "diff-check": ["git", "diff", "--check"],
    }
    results = read("results.json")
    assert len(results) == len(expected_commands)
    assert {row["name"] for row in results} == set(expected_commands)
    for row in results:
        assert row["command"] == expected_commands[row["name"]]
        assert row["exit_code"] == 0, row["name"]
        assert digest(HERE / (row["name"] + ".log")) == row["log_sha256"]
    log = (HERE / "ods-tests.log").read_text()
    summary = r"test result: ok\. (\d+) passed; (\d+) failed; (\d+) ignored"
    counts = re.findall(summary, log)
    assert counts, "test summaries missing"
    totals = [sum(int(row[index]) for row in counts) for index in range(3)]
    assert totals == [1592, 0, 0], totals
    focused = {}
    for target, expected in {"evaluation": 13, "limits": 12, "native": 2, "oracle": 26}.items():
        marker = f"Running tests/ods_formula_text_{target}.rs"
        assert log.count(marker) == 1, marker
        section = log.split(marker, 1)[1].split("Running ", 1)[0]
        rows = re.findall(summary, section)
        assert len(rows) == 1 and int(rows[0][0]) == expected, target
        assert rows[0][1:] == ("0", "0"), target
        focused[target] = int(rows[0][0])
    for test in (
        "exhaustive_small_dyadics_match_integer_reference",
        "reported_boundary_regressions_use_exact_distances",
        "zero_one_subnormal_and_carry_are_stable",
        "invalid_inputs_are_formula_number_errors",
        "work_failure_propagates_before_later_iterations",
    ):
        assert f"text::fraction::tests::{test} ... ok" in log, test
    assert read("verification.json") == {
        "stable_sources": True, "all_required_checks_passed": True,
    }
    print(json.dumps({"verified": True, "passed": totals[0], "failed": 0,
                      "ignored": 0, "gates": 7, "focused": focused}, sort_keys=True))


if __name__ == "__main__":
    main()
