#!/usr/bin/env python3
"""Verify the frozen expression grammar and its retained evidence."""
import csv
import hashlib
import json
from pathlib import Path
import re
import runpy
import shlex
import subprocess
import zipfile

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]


def load(path):
    return json.loads(path.read_text())


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def verify_gates():
    receipt = load(HERE / "gates/results.json")
    assert receipt["sources_unchanged"]
    assert receipt["source_before"] == receipt["source_after"]
    essential = {"crates/litchi-ods/src/codec/formula.rs", "crates/litchi-ods/src/codec/formula/expression.rs",
                 "crates/litchi-ods/src/codec/formula/expression/names.rs", "crates/litchi-ods/src/codec/formula/reference.rs",
                 "crates/litchi-ods/src/codec/formula/reference/iri.rs", "crates/litchi-ods/tests/ods_formula_expressions.rs"}
    assert essential <= set(receipt["source_after"])
    for name, expected in receipt["source_after"].items():
        assert sha(ROOT / name) == expected, name
    assert [item["name"] for item in receipt["commands"]] == ["test", "clippy", "doc", "doctest", "fmt"]
    expected_commands = [
        ["cargo", "test", "--locked", "--offline", "-p", "litchi-ods", "--all-features", "--all-targets"],
        ["cargo", "clippy", "--locked", "--offline", "-p", "litchi-ods", "--all-features", "--all-targets", "--", "-D", "warnings"],
        ["cargo", "doc", "--locked", "--offline", "-p", "litchi-ods", "--all-features", "--no-deps"],
        ["cargo", "test", "--locked", "--offline", "-p", "litchi-ods", "--all-features", "--doc"],
        ["cargo", "fmt", "-p", "litchi-ods", "--", "--check"],
    ]
    assert [item["command"] for item in receipt["commands"]] == expected_commands
    assert all(item["status"] == 0 for item in receipt["commands"])
    assert receipt["environment"]["RUSTDOCFLAGS"] == "-D warnings"
    counts = re.findall(r"test result: ok\. (\d+) passed; (\d+) failed; (\d+) ignored; \d+ measured; 0 filtered out", (HERE / "gates/test.log").read_text())
    requirements = load(HERE / "requirements.json")
    assert len(counts) == requirements["targets"]
    assert sum(int(row[0]) for row in counts) == requirements["tests"]
    assert all(row[1:] == ("0", "0") for row in counts)
    log = (HERE / "gates/test.log").read_text()
    child_summaries = re.findall(r"test result: ok\. (\d+) passed; 0 failed; 0 ignored; 0 measured; ([1-9]\d*) filtered out", log)
    assert len(child_summaries) == requirements["subprocess_tests"] == 1
    assert child_summaries[0] == ("1", "13")
    for name in requirements["required_integration_groups"]:
        assert f"test {name} ... ok" in log, name
    doc = (HERE / "gates/doctest.log").read_text()
    assert f'test result: ok. {requirements["doctests"]} passed; 0 failed; 0 ignored' in doc
    boundaries = load(HERE / "gates/boundaries.json")
    assert boundaries["status"] == 0
    for name, expected in boundaries["files"].items():
        assert sha(ROOT / name) == expected
    assert "crate boundaries valid for" in (HERE / "gates/boundaries.log").read_text()
    return receipt, requirements


def verify_specification():
    spec = load(HERE / "specification.json")
    with zipfile.ZipFile(ROOT / spec["source"]) as archive:
        assert hashlib.sha256(archive.read(spec["entry"])).hexdigest() == spec["entry_sha256"]
    names = load(HERE / "names-character-classes.json")
    assert sha(ROOT / names["helper"]["path"]) == names["helper"]["sha256"]
    comparison = names["independent_expanded_set_check"]
    assert comparison["exact_expanded_set_match"]
    assert comparison["helper_ranges_sorted_and_disjoint"]
    assert comparison["source_code_point_counts"] == {"Letter (BaseChar + Ideographic)": 34514, "CombiningChar": 437, "Digit": 149}
    fixture = load(HERE / "names-character-ranges.json")
    assert fixture["source_sha256"] == names["primary_source"]["sha256"]
    source = (ROOT / names["helper"]["path"]).read_text()
    tables = {"LETTER_RANGES": ["BaseChar", "Ideographic"],
              "COMBINING_RANGES": ["CombiningChar"], "DIGIT_RANGES": ["Digit"]}
    def points(ranges):
        return {value for start, end in ranges for value in range(start, end + 1)}
    for table, productions in tables.items():
        start = source.index("const " + table + ":")
        end = source.index("\n];", start)
        actual = [(int(a, 16), int(b, 16)) for a, b in re.findall(
            r"\(0x([0-9A-Fa-f]+),\s*0x([0-9A-Fa-f]+)\)", source[start:end])]
        expected = [interval for production in productions for interval in fixture["productions"][production]]
        assert points(actual) == points(expected), table
        assert all(a <= b for a, b in actual)
        assert all(left[1] < right[0] for left, right in zip(actual, actual[1:]))


def verify_profiles():
    perf = HERE / "performance"
    legacy = runpy.run_path(str(perf / "harness/run.py"))
    expression = runpy.run_path(str(perf / "expression-harness/run.py"))
    groups = [
        ("baseline", "comparable", "parse", legacy["COMPARABLE_CASES"]),
        ("baseline-literals", "literals", "parse", legacy["LITERAL_CASES"]),
        ("candidate", "comparable", "parse", legacy["COMPARABLE_CASES"]),
        ("candidate-literals", "literals", "parse", legacy["LITERAL_CASES"]),
        ("expression", "expression", "expression", expression["CASES"]),
    ]
    results = {}
    for directory, group, workload, cases in groups:
        folder = perf / directory
        rows = list(csv.DictReader((folder / "raw.csv").open()))
        assert len(rows) == len(cases), directory
        assert [(row["case"], int(row["repeat"])) for row in rows] == list(cases)
        metadata = load(folder / "group.json")
        assert metadata["case_count"] == len(cases)
        assert metadata["cases"] == [{"case": case, "repeat": repeat} for case, repeat in cases]
        commands = (folder / "commands.txt").read_text().splitlines()
        assert len(commands) == len(cases)
        for row, command in zip(rows, commands):
            stem = folder / ("parse-" + row["case"])
            lines = stem.with_suffix(".stdout").read_text().splitlines()
            assert len(lines) == 2 and lines[0].startswith("config ") and lines[1].startswith("result ")
            config, result = [dict(token.split("=", 1) for token in line.split()[1:]) for line in lines]
            for key, value in (config | result).items():
                assert row[key] == value, (directory, row["case"], key)
            assert row["group"] == group and row["workload"] == workload
            assert row["status"] == stem.with_suffix(".status").read_text().strip() == "0"
            assert not stem.with_suffix(".stderr").read_text()
            timing = stem.with_suffix(".time").read_text()
            assert re.search(r"Exit status:\s*0\s*$", timing)
            assert row["max_rss_kib"] == re.search(r"Maximum resident set size \(kbytes\):\s*(\d+)", timing)[1]
            assert int(row["max_rss_kib"]) > 0
            assert int(row["warmups"]) == metadata["warmups"] == 3
            assert int(row["iterations"]) == metadata["iterations"] == 15
            assert row["expected_success"] in {"true", "false"}
            expected = int(row["repeat"]) if row["expected_success"] == "true" else 0
            assert int(row["successes_p50"]) == int(row["successes_max"]) == expected
            assert row["live_before_p50"] == row["live_after_p50"] == row["live_after_max"]
            assert 0 < int(row["p50_ns"]) <= int(row["p95_ns"]) <= int(row["p99_ns"])
            argv = shlex.split(command)
            assert argv[:6] == ["taskset", "-c", "2", "/usr/bin/time", "-v", "-o"]
            assert argv[8:] == ["--workload", workload, "--case", row["case"], "--warmups", "3", "--iterations", "15", "--repeat", row["repeat"]]
        results[directory] = rows
    for before, after in [("baseline", "candidate"), ("baseline-literals", "candidate-literals")]:
        for left, right in zip(results[before], results[after]):
            for key in ("case", "input_bytes", "repeat", "expected_success", "checksum_p50", "checksum_max", "successes_p50", "alloc_calls_p50", "requested_bytes_p50", "peak_live_delta_p50"):
                assert left[key] == right[key], (left["case"], key)
    binaries = load(HERE / "gates/root-binary-verification.json")["binaries"]
    for sequence_name, directories in [("baseline-sequence.json", ["baseline", "baseline-literals"]),
                                       ("candidate-sequence.json", ["candidate", "candidate-literals", "expression"])]:
        sequence = load(perf / sequence_name)
        assert len(sequence["jobs"]) == len(directories)
        for job, directory in zip(sequence["jobs"], directories):
            label = "baseline" if directory.startswith("baseline") else "expression" if directory == "expression" else "candidate"
            assert job["binary_sha256"] == binaries[label]["sha256"]
            assert job["raw_sha256"] == sha(perf / directory / "raw.csv")
            assert job["row_count"] == len(results[directory]) and job["status"] == 0
    return sum(len(rows) for rows in results.values())


def verify_sources(receipt):
    replay = load(HERE / "gates/patch-replay.json")
    assert replay["status"] == 0 and replay["source_after"] == receipt["source_after"]
    assert sha(HERE / "candidate.patch") == replay["patch_sha256"]
    # Replay each unified hunk in memory against the immutable baseline.
    # This remains checkable after the disposable worktree and ELFs are removed.
    base = load(HERE / "specification.json")["base_commit"]
    patch = (HERE / "candidate.patch").read_text()
    sections = re.split(r"(?m)^diff --git a/(.+) b/(.+)\n", patch)
    assert sections[0] == "" and (len(sections) - 1) % 3 == 0
    changed = set()
    for index in range(1, len(sections), 3):
        left, name, body = sections[index:index + 3]
        assert left == name and name not in changed
        assert not Path(name).is_absolute() and ".." not in Path(name).parts
        changed.add(name)
        if "--- /dev/null\n" in body:
            assert subprocess.run(["git", "cat-file", "-e", base + ":" + name], cwd=ROOT, stderr=subprocess.DEVNULL).returncode != 0
            old = []
        else:
            old = subprocess.check_output(["git", "show", base + ":" + name], cwd=ROOT).decode().splitlines(keepends=True)
        output, cursor = [], 0
        hunks = re.split(r"(?m)^@@ -(\d+)(?:,(\d+))? \+(\d+)(?:,(\d+))? @@[^\n]*\n", body)
        assert len(hunks) > 1 and (len(hunks) - 1) % 5 == 0
        for h in range(1, len(hunks), 5):
            old_start, old_count, new_start, new_count, content = hunks[h:h + 5]
            start = max(0, int(old_start) - 1)
            assert start >= cursor
            output.extend(old[cursor:start]); cursor = start
            assert len(output) == max(0, int(new_start) - 1)
            removed = added = 0
            for line in content.splitlines(keepends=True):
                assert line[0] in " +-", (name, line)
                if line[0] in " -":
                    assert cursor < len(old) and old[cursor] == line[1:], name
                    cursor += 1; removed += 1
                if line[0] in " +":
                    output.append(line[1:]); added += 1
            assert removed == int(old_count or 1) and added == int(new_count or 1)
        output.extend(old[cursor:])
        assert "".join(output).encode() == (ROOT / name).read_bytes(), name
    assert changed == set(replay["changed_files"])
    baseline = load(HERE / "performance/baseline/source-sha256.json")
    assert baseline["source_before"] == baseline["source_after"]
    for name, expected in baseline["source_before"].items():
        content = subprocess.check_output(["git", "show", baseline["commit"] + ":" + name], cwd=ROOT)
        assert hashlib.sha256(content).hexdigest() == expected, name
    binaries = load(HERE / "gates/root-binary-verification.json")["binaries"]
    assert set(binaries) == {"baseline", "candidate", "expression"}
    for label, verified in binaries.items():
        folder = HERE / "performance" / label
        provenance = load(folder / "binary-provenance.json")
        assert provenance["sha256"] == verified["sha256"]
        assert provenance["size_bytes"] == verified["bytes"] > 0
        assert provenance["build_status"] == 0
        assert (folder / "binary-sha256.txt").read_text().split()[0] == verified["sha256"]
        harness = HERE / "performance" / ("expression-harness" if label == "expression" else "harness")
        harness_entries = {}
        for line in (folder / "harness-sha256.txt").read_text().splitlines():
            expected, name = line.split("  ", 1)
            assert name in {"Cargo.toml", "Cargo.lock", "run.py", "src/main.rs"}
            assert name not in harness_entries and sha(harness / name) == expected
            harness_entries[name] = expected
        assert len(harness_entries) == 4
        assert load(folder / "build-command.json")["status"] == 0
        assert (folder / "build.status").read_text().strip() == "0"
        if label != "baseline":
            sources = load(folder / "source-sha256.json")
            assert sources["source_before"] == sources["source_after"] == receipt["source_after"]


def verify_counters():
    folder = HERE / "performance/ast-perf-stat"
    sequence = load(folder / "sequence.json")
    digest = load(HERE / "gates/root-binary-verification.json")["binaries"]["expression"]["sha256"]
    assert sequence["binary_sha256"] == digest
    assert sequence["cpu_pin"] == "2"
    assert [job["case"] for job in sequence["jobs"]] == ["expr-flat-4096", "expr-name-4096", "expr-reference-4096"]
    assert (folder / "commands.txt").read_text().splitlines() == [job["command"] for job in sequence["jobs"]]
    for job in sequence["jobs"]:
        assert job["binary_sha256"] == digest and job["status"] == 0
        assert job["elapsed_seconds"] > 0 and job["finished_at"] >= job["started_at"]
        for field in ("stdout", "stderr", "counters"):
            assert sha(folder / job[field]) == job[field + "_sha256"]
        assert not (folder / job["stderr"]).read_text()
        assert (folder / ("ast-" + job["label"] + ".status")).read_text().strip() == "0"
        lines = (folder / job["stdout"]).read_text().splitlines()
        config, result = [dict(token.split("=", 1) for token in line.split()[1:]) for line in lines]
        assert config["case"] == job["case"] and config["workload"] == "expression"
        assert config["repeat"] == result["successes_p50"] == result["successes_max"] == "8"
        assert config["warmups"] == "3" and config["iterations"] == "15"
        assert result["live_before_p50"] == result["live_after_p50"] == result["live_after_max"]
        counters = [row for row in csv.reader((folder / job["counters"]).open()) if row and not row[0].startswith("#")]
        assert [row[2] for row in counters] == sequence["events"] == ["cycles", "instructions", "branches", "branch-misses", "cache-misses"]
        assert all(float(row[0]) > 0 and float(row[4]) > 0 for row in counters)
    return len(sequence["jobs"])


def verify_manifest():
    entries = {}
    for line in (HERE / "artifacts.sha256").read_text().splitlines():
        expected, name = line.split("  ", 1)
        path = Path(name)
        assert not path.is_absolute() and ".." not in path.parts
        assert name not in entries
        assert sha(HERE / name) == expected, name
        entries[name] = expected
    actual = {str(path.relative_to(HERE)) for path in HERE.rglob("*")
              if path.is_file() and path.name not in {"artifacts.sha256", "root-verification.json"}
              and "__pycache__" not in path.parts}
    assert set(entries) == actual
    return len(entries)


def main():
    receipt, requirements = verify_gates()
    verify_specification()
    verify_sources(receipt)
    lanes = verify_profiles()
    counters = verify_counters()
    count = verify_manifest()
    print(json.dumps({"verified": True, "tests": requirements["tests"],
                      "targets": requirements["targets"], "profile_lanes": lanes, "counter_captures": counters, "artifact_count": count}))


if __name__ == "__main__":
    main()
