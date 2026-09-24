#!/usr/bin/env python3
"""Verify source-bound formula-string gates and retained measurements.

The benchmark receipts contain several archived source states. This verifier
checks each state against the receipt which produced it, then checks raw lane
output and the per-lane ``/usr/bin/time`` RSS record. The global artifact
manifest is mandatory and covers the complete retained evidence bundle.
"""

import csv
import hashlib
import json
from pathlib import Path
import re
import subprocess
import zipfile


HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]
PERF = HERE / "performance"
HARNESS = PERF / "harness"

BASE_COMMIT = "cbc60f1123d105f67f4abdd0d53034fb3325c25f"
FORMULA = "crates/litchi-ods/src/codec/formula.rs"
STRING_TEST = "crates/litchi-ods/tests/ods_formula_string_regression.rs"
EXPECTED_GROUPS = {"comparable": 12, "coverage": 16, "literals": 9}
FINAL_DIRS = [
    ("baseline", "comparable"),
    ("candidate", "comparable"),
    ("baseline-coverage", "coverage"),
    ("candidate-coverage", "coverage"),
    ("baseline-literals", "literals"),
    ("candidate-literals", "literals"),
]
PRE_NUL_DIRS = [
    ("baseline-pre-nul", "comparable"),
    ("candidate-pre-nul", "comparable"),
    ("baseline-coverage-pre-nul", "coverage"),
    ("candidate-coverage-pre-nul", "coverage"),
    ("baseline-literals-pre-nul", "literals"),
    ("candidate-literals-pre-nul", "literals"),
]
CONFIG_FIELDS = {
    "workload",
    "case",
    "input_bytes",
    "repeat",
    "warmups",
    "iterations",
    "expected_success",
}
RESULT_FIELDS = {
    "mean_ns",
    "p50_ns",
    "p95_ns",
    "p99_ns",
    "alloc_calls_p50",
    "alloc_calls_max",
    "dealloc_calls_p50",
    "dealloc_calls_max",
    "requested_bytes_p50",
    "requested_bytes_max",
    "released_bytes_p50",
    "released_bytes_max",
    "live_before_p50",
    "live_after_p50",
    "live_after_max",
    "peak_live_delta_p50",
    "peak_live_delta_max",
    "successes_p50",
    "successes_max",
    "checksum_p50",
    "checksum_max",
}
NUMERIC_FIELDS = RESULT_FIELDS | {
    "input_bytes",
    "repeat",
    "warmups",
    "iterations",
    "max_rss_kib",
    "status",
}
EVENTS = ["cycles", "instructions", "branches", "branch-misses", "cache-misses"]
RAW_FIELDS = [
    "group",
    "workload",
    "case",
    "input_bytes",
    "repeat",
    "warmups",
    "iterations",
    "expected_success",
    "mean_ns",
    "p50_ns",
    "p95_ns",
    "p99_ns",
    "alloc_calls_p50",
    "alloc_calls_max",
    "dealloc_calls_p50",
    "dealloc_calls_max",
    "requested_bytes_p50",
    "requested_bytes_max",
    "released_bytes_p50",
    "released_bytes_max",
    "live_before_p50",
    "live_after_p50",
    "live_after_max",
    "peak_live_delta_p50",
    "peak_live_delta_max",
    "successes_p50",
    "successes_max",
    "checksum_p50",
    "checksum_max",
    "max_rss_kib",
    "status",
]


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def load(path):
    return json.loads(path.read_text())


def assert_file_hash(path, expected):
    assert path.is_file(), path
    actual = sha(path)
    assert actual == expected, (path, expected, actual)


def parse_hash_lines(path, base, verify_files=True):
    """Return and verify a conventional ``sha256  relative/path`` manifest."""
    result = {}
    for line_number, line in enumerate(path.read_text().splitlines(), 1):
        if not line.strip():
            continue
        fields = line.split("  ", 1)
        assert len(fields) == 2, (path, line_number, line)
        expected, name = fields
        assert re.fullmatch(r"[0-9a-f]{64}", expected), (path, line_number)
        relative = Path(name)
        assert not relative.is_absolute() and ".." not in relative.parts, (
            path,
            line_number,
            name,
        )
        if verify_files:
            target = base / relative
            assert_file_hash(target, expected)
        assert name not in result, (path, name)
        result[name] = expected
    assert result, path
    return result


def validate_source_text_manifest(path, expected, verify_files=False):
    assert parse_hash_lines(path, ROOT, verify_files) == expected, path


def parse_result_summaries(path):
    pattern = re.compile(
        r"test result: (ok|FAILED)\. (\d+) passed; (\d+) failed; "
        r"(\d+) ignored;"
    )
    return [
        (state, int(passed), int(failed), int(ignored))
        for state, passed, failed, ignored in pattern.findall(path.read_text())
    ]


def validate_gate_log(directory, expected_tests):
    summaries = parse_result_summaries(directory / "test.log")
    assert len(summaries) == expected_tests["targets"], directory
    assert sum(item[1] for item in summaries) == expected_tests["tests"], directory
    assert sum(item[2] for item in summaries) == expected_tests["failed"], directory
    assert sum(item[3] for item in summaries) == expected_tests["ignored"], directory
    assert all(
        state == "ok" and failed == 0 and ignored == 0
        for state, _, failed, ignored in summaries
    ), directory
    for name in ["test", "clippy", "doc", "doctest", "fmt"]:
        assert (directory / (name + ".log")).is_file()


def validate_gate_receipt(directory, expected_sources, expected_tests):
    receipt = load(directory / "results.json")
    assert receipt["head"] == BASE_COMMIT
    assert receipt["sources_unchanged"]
    assert receipt["source_before"] == receipt["source_after"] == expected_sources
    assert len(receipt["commands"]) == 5
    assert [item["name"] for item in receipt["commands"]] == [
        "test",
        "clippy",
        "doc",
        "doctest",
        "fmt",
    ]
    assert all(item["status"] == 0 for item in receipt["commands"])
    assert receipt["environment"] == {
        "CARGO_TARGET_DIR": "/var/tmp/ods-string-target",
        "TMPDIR": "/var/tmp/ods-string-tmp",
        "RUSTDOCFLAGS": "-D warnings",
    }
    validate_gate_log(directory, expected_tests)


def validate_patch_replay(path, receipt_path, expected_sources):
    receipt = load(receipt_path)
    assert receipt["base"] == BASE_COMMIT
    assert receipt["verified"]
    assert receipt["patch_sha256"] == sha(path)
    assert set(receipt["source_after"]) == {FORMULA, STRING_TEST}
    assert receipt["source_after"] == {
        name: expected_sources[name] for name in (FORMULA, STRING_TEST)
    }


def validate_source_json(path, expected, commit_key, verify_files=False):
    receipt = load(path)
    assert receipt[commit_key] == BASE_COMMIT
    if "source_before" in receipt:
        assert receipt["source_before"] == receipt["source_after"]
    assert receipt["source_after"] == expected
    validate_source_text_manifest(
        path.with_name("source-sha256.txt"), expected, verify_files
    )


def validate_baseline_negative(directory, baseline_formula, string_test):
    source_manifest = load(directory / "isolated-string-regression-source-sha256.json")
    expected = {FORMULA: baseline_formula, STRING_TEST: string_test}
    assert source_manifest["base_commit"] == BASE_COMMIT
    assert source_manifest["source_before"] == source_manifest["source_after"] == expected
    assert source_manifest["formula_source_role"] == "baseline"
    assert source_manifest["test_fixture_source_role"] == "candidate-regression-fixture"

    receipt = load(directory / "isolated-string-regression.json")
    assert receipt["status"] == receipt["expected_baseline_status"] == 101
    assert receipt["expected_passes"] == 4
    expected_failures = {
        "many_short_and_empty_utf8_literals_retain_bounded_capacity",
        "nul_is_rejected_while_non_nul_controls_remain_valid",
    }
    assert set(receipt["expected_failures"]) == expected_failures
    assert receipt["log"] == "baseline/isolated-string-regression.log"
    assert receipt["source_manifest"] == "baseline/isolated-string-regression-source-sha256.json"
    assert receipt["source_manifest_sha256"] == sha(
        PERF / "baseline/isolated-string-regression-source-sha256.json"
    )
    assert (directory / "isolated-string-regression.status").read_text().strip() == "101"
    log = PERF / "baseline/isolated-string-regression.log"
    assert re.search(r"running 6 tests\b", log.read_text())
    assert parse_result_summaries(log) == [("FAILED", 4, 2, 0)]
    failed_names = set(
        re.findall(r"^test ([^ ]+) \.\.\. FAILED$", log.read_text(), re.MULTILINE)
    )
    assert failed_names == expected_failures, failed_names


def validate_final_isolated_checks(candidate_sources):
    receipt = load(PERF / "candidate/isolated-checks.json")
    assert receipt["source_manifest"] == "candidate/source-sha256.json"
    assert receipt["total_tests"] == 82
    expected = {
        "formula-unit": 49,
        "string-regression": 6,
        "scan-regression": 6,
        "tokenizer-regression": 4,
        "function-catalog": 7,
        "reference-integration": 10,
    }
    assert {check["name"]: check["expected_passes"] for check in receipt["checks"]} == expected
    assert sum(expected.values()) == receipt["total_tests"]
    for check in receipt["checks"]:
        assert check["status"] == 0
        assert check["observed_passes"] == expected[check["name"]]
        log = PERF / check["command_log"]
        assert log.is_file() and str(log).startswith(str(PERF / "candidate"))
        assert parse_result_summaries(log) == [("ok", expected[check["name"]], 0, 0)]
        status = PERF / "candidate" / ("isolated-" + check["name"] + ".status")
        assert status.read_text().strip() == "0"
    assert candidate_sources["source_after"] == load(
        PERF / "candidate/source-sha256.json"
    )["source_after"]


def validate_harness_manifests(directory_names):
    expected = None
    for name in directory_names:
        actual = parse_hash_lines(PERF / name / "harness-sha256.txt", HARNESS)
        if expected is None:
            expected = actual
        else:
            assert actual == expected, name
    assert expected is not None


def validate_binary_receipts():
    root_receipt = load(HERE / "gates/root-binary-verification.json")
    binaries = root_receipt["binaries"]
    assert set(binaries) == {"baseline", "candidate-initial", "candidate"}
    directories = {
        "baseline": "baseline",
        "candidate-initial": "candidate-pre-nul",
        "candidate": "candidate",
    }
    expected = {}
    for kind, directory in directories.items():
        digest = (PERF / directory / "binary-sha256.txt").read_text().split()[0]
        expected[kind] = digest
        receipt = binaries[kind]
        assert receipt["sha256"] == digest
        assert receipt["bytes"] > 0
        assert receipt["captured_path"]
        provenance = load(PERF / directory / "binary-provenance.json")
        assert provenance["sha256"] == digest
        assert provenance["size_bytes"] == receipt["bytes"]
        assert provenance["base_commit"] == BASE_COMMIT
        assert provenance["build_status"] == 0
        assert provenance["source_manifest"] == directory + "/source-sha256.json"
        build = load(PERF / directory / "build-command.json")
        assert build["status"] == 0
        assert build["CARGO_TARGET_DIR"] == "/var/tmp/ods-string-target"
        assert build["TMPDIR"] == "/var/tmp/ods-string-tmp"
        # The measured executables are intentionally removed after capture;
        # when an archive retains one, validate it too.
        binary = PERF / directory / "ods-formula-reference-profile"
        if binary.exists():
            assert_file_hash(binary, digest)
            assert binary.stat().st_size == receipt["bytes"]

    # The pre-NUL binary was archived under candidate-pre-nul after the target
    # directory was reused; this maps candidate-initial to that archive.
    assert expected["candidate-initial"] == (
        PERF / "candidate-pre-nul/binary-sha256.txt"
    ).read_text().split()[0]


def parse_kv_line(line, prefix, path):
    assert line.startswith(prefix), (path, line)
    values = {}
    for field in line[len(prefix) :].split():
        key, separator, value = field.partition("=")
        assert separator and key and key not in values, (path, line)
        values[key] = value
    return values


def validate_lane_files(folder, row):
    stem = folder / ("parse-" + row["case"])
    stdout = stem.with_suffix(".stdout")
    lines = [line for line in stdout.read_text().splitlines() if line]
    assert len(lines) == 2, stdout
    config = parse_kv_line(lines[0], "config ", stdout)
    result = parse_kv_line(lines[1], "result ", stdout)
    assert set(config) == CONFIG_FIELDS, (stdout, config)
    assert set(result) == RESULT_FIELDS, (stdout, result)
    for key, value in config.items():
        assert row[key] == value, (stdout, key)
    for key, value in result.items():
        assert row[key] == value, (stdout, key)

    assert stem.with_suffix(".status").read_text().strip() == "0"
    assert row["status"] == "0"
    assert stem.with_suffix(".stderr").read_text() == ""
    time_text = stem.with_suffix(".time").read_text()
    rss = re.search(r"Maximum resident set size \(kbytes\):\s*(\d+)", time_text)
    assert rss, stem.with_suffix(".time")
    assert row["max_rss_kib"] == rss.group(1), (row["case"], row["max_rss_kib"])
    assert re.search(r"^\s*Exit status:\s*0\s*$", time_text, re.MULTILINE)

    for key in NUMERIC_FIELDS:
        assert key in row, (folder, row["case"], key)
        assert int(row[key]) >= 0, (folder, row["case"], key)
    repeat = int(row["repeat"])
    expected_successes = repeat if row["expected_success"] == "true" else 0
    assert int(row["successes_p50"]) == int(row["successes_max"]) == expected_successes
    assert row["warmups"] == "3" and row["iterations"] == "15"


def validate_perf_folder(name, group, expected_binary):
    folder = PERF / name
    expected_count = EXPECTED_GROUPS[group]
    metadata = load(folder / "group.json")
    assert metadata["group"] == group
    assert metadata["case_count"] == expected_count
    assert metadata["warmups"] == 3 and metadata["iterations"] == 15
    expected_cases = [(item["case"], item["repeat"]) for item in metadata["cases"]]
    assert len(expected_cases) == expected_count

    with (folder / "raw.csv").open(newline="") as stream:
        rows = list(csv.DictReader(stream))
    assert len(rows) == expected_count, name
    assert list(rows[0]) == RAW_FIELDS, name
    assert [(row["case"], int(row["repeat"])) for row in rows] == expected_cases

    commands = [line for line in (folder / "commands.txt").read_text().splitlines() if line]
    assert len(commands) == expected_count, (name, len(commands))
    for row, command in zip(rows, commands):
        assert f"--case {row['case']}" in command
        assert f"--repeat {row['repeat']}" in command
        assert "--warmups 3" in command and "--iterations 15" in command
        assert "/usr/bin/time -v" in command

    for row in rows:
        assert row["group"] == group and row["workload"] == "parse"
        validate_lane_files(folder, row)

    run_end = load(folder / "run-end.json")
    assert run_end["status"] == 0
    assert run_end["group"] == group
    assert run_end["row_count"] == expected_count
    assert run_end["binary_sha256"] == expected_binary
    assert run_end["warmups"] == 3 and run_end["iterations"] == 15
    assert run_end["raw_sha256"] == sha(folder / "raw.csv")
    return rows


def validate_comparability(rows, left_name, right_name):
    assert len(rows[left_name]) == len(rows[right_name])
    for left, right in zip(rows[left_name], rows[right_name]):
        for key in ["case", "input_bytes", "repeat", "expected_success"]:
            assert left[key] == right[key], (left_name, right_name, key)
        if left["expected_success"] == "true":
            assert left["checksum_p50"] == right["checksum_p50"]
            assert left["checksum_max"] == right["checksum_max"]


def sequence_folder(label, suffix=""):
    return {
        "baseline-comparable": "baseline" + suffix,
        "candidate-comparable": "candidate" + suffix,
        "baseline-coverage": "baseline-coverage" + suffix,
        "candidate-coverage": "candidate-coverage" + suffix,
        "baseline-literals": "baseline-literals" + suffix,
        "candidate-literals": "candidate-literals" + suffix,
    }[label]


def validate_sequence(path, suffix, binary_by_side):
    sequence = load(path)
    assert sequence["warmups"] == 3 and sequence["iterations"] == 15
    if suffix:
        labels = [
            "baseline-comparable",
            "candidate-comparable",
            "baseline-coverage",
            "candidate-coverage",
            "baseline-literals",
            "candidate-literals",
        ]
    else:
        labels = [
            "baseline-comparable",
            "baseline-coverage",
            "baseline-literals",
            "candidate-comparable",
            "candidate-coverage",
            "candidate-literals",
        ]
    assert [job["label"] for job in sequence["jobs"]] == labels
    for job in sequence["jobs"]:
        folder = sequence_folder(job["label"], suffix)
        side = job["label"].split("-", 1)[0]
        assert job["status"] == 0
        assert job["group"] in EXPECTED_GROUPS
        assert job["binary_sha256"] == binary_by_side[side]
        assert job["row_count"] == EXPECTED_GROUPS[job["group"]]
        assert job["raw_sha256"] == sha(PERF / folder / "raw.csv")
    return sequence


def validate_paired_raw_manifest(path, suffix, sequence):
    expected = {}
    for job in sequence["jobs"]:
        expected[sequence_folder(job["label"], suffix) + "/raw.csv"] = job["raw_sha256"]
    assert parse_hash_lines(path, PERF) == expected


def validate_targeted_abab(binary_by_side):
    directory = PERF / "targeted-abab"
    sequence = load(directory / "sequence.json")
    assert sequence["harness_sha256"] == sha(HARNESS / "run.py")
    assert sequence["warmups"] == 3 and sequence["iterations"] == 15
    assert sequence["order"] == "A/B/A/B for each case (baseline,candidate,baseline,candidate)"
    cases = sequence["cases"]
    assert cases == [
        "literal-many-short-64",
        "literal-many-short-256",
        "literal-many-short-1024",
        "literal-single-plain-64k",
        "literal-single-doubled-quote-64k",
    ]
    jobs = sequence["jobs"]
    assert len(jobs) == len(cases) * 4
    expected_manifest = {}
    for offset, case in enumerate(cases):
        for index, (round_number, side) in enumerate(
            [(1, "baseline"), (2, "candidate"), (3, "baseline"), (4, "candidate")]
        ):
            job = jobs[offset * 4 + index]
            assert job["case"] == case and job["round"] == round_number
            assert job["side"] == side and job["status"] == 0
            assert job["binary_sha256"] == binary_by_side[side]
            raw = PERF / job["raw"]
            assert raw.is_file()
            assert job["raw_sha256"] == sha(raw)
            expected_manifest[str(raw.relative_to(PERF))] = job["raw_sha256"]
            with raw.open(newline="") as stream:
                rows = list(csv.DictReader(stream))
            assert len(rows) == 1 and rows[0]["case"] == case
            assert int(rows[0]["repeat"]) == int(job["repeat"])
            validate_lane_files(raw.parent, rows[0])
    assert parse_hash_lines(directory / "raw-sha256.txt", PERF) == expected_manifest


def validate_hardware_perf_stat(binary_by_side):
    directory = PERF / "hardware" / "perf-stat"
    summary = load(directory / "summary.json")
    assert summary["events"] == EVENTS
    assert summary["warmups"] == 3 and summary["iterations"] == 15
    rows = summary["rows"]
    expected_cases = {
        ("baseline", "literal-many-short-4096"),
        ("candidate", "literal-many-short-4096"),
        ("baseline", "literal-single-doubled-quote-64k"),
        ("candidate", "literal-single-doubled-quote-64k"),
    }
    assert len(rows) == 4 and {(r["side"], r["case"]) for r in rows} == expected_cases
    for row in rows:
        assert row["status"] == 0
        assert row["binary_sha256"] == binary_by_side[row["side"]]
        assert set(row["counters"]) == set(EVENTS)
        assert all(int(row["counters"][event]) > 0 for event in EVENTS)
        stdout = PERF / row["stdout"]
        stderr = PERF / row["stderr"]
        command = PERF / row["command"]
        assert stdout.is_file() and stderr.is_file() and command.is_file()
        status = directory / f"{row['side']}-{row['case']}.status"
        assert status.read_text().strip() == "0"
        output = [line for line in stdout.read_text().splitlines() if line]
        assert len(output) == 2
        config = parse_kv_line(output[0], "config ", stdout)
        result = parse_kv_line(output[1], "result ", stdout)
        assert config["workload"] == "parse" and config["case"] == row["case"]
        assert int(config["repeat"]) == row["repeat"]
        assert config["warmups"] == "3" and config["iterations"] == "15"
        assert config["expected_success"] == "true"
        assert set(result) == RESULT_FIELDS
        assert int(result["successes_p50"]) == int(result["successes_max"]) == row["repeat"]
        counters = list(csv.reader(stderr.read_text().splitlines()))
        assert len(counters) == len(EVENTS)
        assert {item[2] for item in counters} == set(EVENTS)
        for item in counters:
            assert len(item) >= 3 and int(item[0].strip()) > 0
            assert item[0].strip() == row["counters"][item[2]]


def validate_global_manifest():
    manifest = HERE / "artifacts.sha256"
    if not manifest.exists():
        raise AssertionError(f"missing mandatory artifact manifest: {manifest}")
    entries = {}
    for line_number, line in enumerate(manifest.read_text().splitlines(), 1):
        if not line.strip():
            continue
        fields = line.split("  ", 1)
        assert len(fields) == 2, (manifest, line_number, line)
        expected, name = fields
        assert re.fullmatch(r"[0-9a-f]{64}", expected), (manifest, line_number)
        relative = Path(name)
        assert not relative.is_absolute() and ".." not in relative.parts
        assert name not in entries, (manifest, name)
        target = HERE / relative
        if target == manifest:
            raise AssertionError("artifact manifest must not hash itself")
        assert_file_hash(target, expected)
        entries[name] = expected
    assert len(entries) >= 300, len(entries)
    actual = {
        str(path.relative_to(HERE)) for path in HERE.rglob("*")
        if path.is_file() and path.name not in {"artifacts.sha256", "root-verification.json"}
        and "__pycache__" not in path.parts
    }
    assert set(entries) == actual, (set(entries) - actual, actual - set(entries))
    return f"{len(entries)} files"


def main():
    specification = load(HERE / "specification.json")
    with zipfile.ZipFile(ROOT / specification["source"]) as archive:
        content = archive.read(specification["entry"])
        assert hashlib.sha256(content).hexdigest() == specification["entry_sha256"]

    final_gates = load(HERE / "gates/results.json")
    initial_gates = load(HERE / "gates/initial/results.json")
    assert final_gates["head"] == initial_gates["head"] == BASE_COMMIT
    final_sources = final_gates["source_after"]
    initial_sources = initial_gates["source_after"]
    expected_tests = load(HERE / "requirements.json")
    validate_gate_receipt(HERE / "gates", final_sources, expected_tests)
    validate_gate_receipt(HERE / "gates/initial", initial_sources, expected_tests)
    for name, expected in final_sources.items():
        assert_file_hash(ROOT / name, expected)

    validate_patch_replay(
        HERE / "initial-candidate.patch",
        HERE / "gates/initial/root-patch-replay.json",
        {FORMULA: initial_sources[FORMULA], STRING_TEST: initial_sources[STRING_TEST]},
    )
    validate_patch_replay(
        HERE / "candidate.patch",
        HERE / "gates/root-patch-replay.json",
        {FORMULA: final_sources[FORMULA], STRING_TEST: final_sources[STRING_TEST]},
    )
    assert (HERE / "candidate-patch-sha256.txt").read_text().split() == [
        sha(HERE / "candidate.patch"),
        "candidate.patch",
    ]

    baseline_sources = load(PERF / "baseline/source-sha256.json")
    assert baseline_sources["commit"] == BASE_COMMIT
    assert baseline_sources["source_before"] == baseline_sources["source_after"]
    for name, expected in baseline_sources["source_after"].items():
        content = subprocess.check_output(["git", "show", BASE_COMMIT + ":" + name], cwd=ROOT)
        assert hashlib.sha256(content).hexdigest() == expected, name
    validate_source_text_manifest(
        PERF / "baseline/source-sha256.txt", baseline_sources["source_after"]
    )

    validate_source_json(
        PERF / "candidate-pre-nul/source-sha256.json",
        initial_sources,
        "base_commit",
    )
    candidate_sources = load(PERF / "candidate/source-sha256.json")
    validate_source_json(
        PERF / "candidate/source-sha256.json", final_sources, "base_commit", True
    )
    assert load(PERF / "candidate-pre-nul/source-sha256.json")["root_gate_source_after_match"]
    assert candidate_sources["root_gate_source_after_match"]

    baseline_formula = baseline_sources["source_after"][FORMULA]
    validate_baseline_negative(PERF / "baseline", baseline_formula, final_sources[STRING_TEST])
    validate_baseline_negative(PERF / "baseline-pre-nul", baseline_formula, final_sources[STRING_TEST])
    validate_final_isolated_checks(candidate_sources)

    validate_harness_manifests(
        ["baseline", "candidate", "baseline-pre-nul", "candidate-pre-nul"]
    )
    validate_binary_receipts()
    binary_by_side = {
        "baseline": (PERF / "baseline/binary-sha256.txt").read_text().split()[0],
        "candidate": (PERF / "candidate/binary-sha256.txt").read_text().split()[0],
    }
    pre_binary_by_side = {
        "baseline": binary_by_side["baseline"],
        "candidate": (PERF / "candidate-pre-nul/binary-sha256.txt").read_text().split()[0],
    }

    all_rows = {}
    for name, group in FINAL_DIRS:
        all_rows[name] = validate_perf_folder(name, group, binary_by_side[name.split("-", 1)[0]])
    for name, group in PRE_NUL_DIRS:
        all_rows[name] = validate_perf_folder(name, group, pre_binary_by_side[name.split("-", 1)[0]])
    for side in ["baseline", "candidate"]:
        for group in ["comparable", "coverage", "literals"]:
            final_name = side + ("" if group == "comparable" else "-" + group)
            pre_name = final_name + "-pre-nul"
            assert [
                (row["case"], row["input_bytes"], row["repeat"], row["expected_success"])
                for row in all_rows[final_name]
            ] == [
                (row["case"], row["input_bytes"], row["repeat"], row["expected_success"])
                for row in all_rows[pre_name]
            ]
    for left, right in [
        ("baseline", "candidate"),
        ("baseline-coverage", "candidate-coverage"),
        ("baseline-literals", "candidate-literals"),
        ("baseline-pre-nul", "candidate-pre-nul"),
        ("baseline-coverage-pre-nul", "candidate-coverage-pre-nul"),
        ("baseline-literals-pre-nul", "candidate-literals-pre-nul"),
    ]:
        validate_comparability(all_rows, left, right)

    final_sequence = validate_sequence(PERF / "paired/sequence.json", "", binary_by_side)
    duplicate_sequence = load(PERF / "paired-sequence.json")
    for key in ["harness", "harness_sha256", "warmups", "iterations", "jobs", "order"]:
        assert duplicate_sequence[key] == final_sequence[key], key
    assert final_sequence["harness_sha256"] == sha(HARNESS / "run.py")
    validate_paired_raw_manifest(PERF / "paired/raw-sha256.txt", "", final_sequence)

    pre_sequence = validate_sequence(
        PERF / "pre-nul-paired-sequence.json", "-pre-nul", pre_binary_by_side
    )
    assert "harness_sha256" not in pre_sequence
    validate_paired_raw_manifest(PERF / "pre-nul-paired-raw-sha256.txt", "-pre-nul", pre_sequence)
    historical = load(PERF / "pre-nul-sequence.json")
    assert historical["candidate_binary_sha256"] == pre_binary_by_side["candidate"]
    assert historical["candidate_source_sha256"] == initial_sources[FORMULA]
    assert historical["status"].startswith("superseded by")
    assert len(historical["entries"]) == 3
    for entry in historical["entries"]:
        archived = PERF / entry["archived_as"] / "raw.csv"
        with archived.open(newline="") as stream:
            assert entry["rows"] == len(list(csv.DictReader(stream)))
        assert entry["raw_sha256"] == sha(archived)

    validate_targeted_abab(binary_by_side)
    validate_hardware_perf_stat(binary_by_side)
    manifest_state = validate_global_manifest()
    print(
        json.dumps(
            {
                "verified": True,
                "tests": expected_tests["tests"],
                "targets": expected_tests["targets"],
                "isolated_tests": 82,
                "performance_lanes": 74,
                "initial_performance_lanes": 74,
                "targeted_abab_lanes": 20,
                "hardware_perf_stat_rows": 4,
                "global_manifest": manifest_state,
            }
        )
    )


if __name__ == "__main__":
    main()
