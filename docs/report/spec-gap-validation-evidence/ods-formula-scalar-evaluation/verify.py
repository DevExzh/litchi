#!/usr/bin/env python3
"""Verify the scalar evaluation batch and its retained evidence."""
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

def verify_gates():
    receipt = load(HERE / "gates/results.json")
    assert receipt["sources_unchanged"]
    assert receipt["source_before"] == receipt["source_after"]
    essential = {"crates/litchi-ods/src/codec/formula/evaluation.rs", "crates/litchi-ods/tests/ods_formula_evaluation.rs",
                 "crates/litchi-ods/src/codec/formula.rs", "crates/litchi-ods/src/codec/formula/expression.rs",
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
    fixture = load(HERE / "scalar-cases.json")
    assert fixture["spec"]["entry_sha256"] == spec["entry_sha256"]
    ids = [case["id"] for group in ["normative_cases", "host_dependent_cases"] for case in fixture[group]]
    assert len(ids) == len(set(ids))
    names = load(HERE / "names-verification.json")
    assert sha(ROOT / names["helper"]) == names["helper_sha256"]
    assert sha(ROOT / names["independent_fixture"]) == names["fixture_sha256"]
    ranges = load(ROOT / names["independent_fixture"])
    source = (ROOT / names["helper"]).read_text()
    def points(intervals):
        return {value for start, end in intervals for value in range(start, end + 1)}
    for table, productions in {"LETTER_RANGES": ["BaseChar", "Ideographic"],
            "DIGIT_RANGES": ["Digit"], "COMBINING_RANGES": ["CombiningChar"]}.items():
        start = source.index("const " + table + ":")
        end = source.index("\n];", start)
        actual = [(int(a, 16), int(b, 16)) for a,b in re.findall(
            r"\(0x([0-9A-Fa-f]+),\s*0x([0-9A-Fa-f]+)\)", source[start:end])]
        expected = [interval for production in productions for interval in ranges["productions"][production]]
        assert points(actual) == points(expected)
        assert len(points(actual)) == names["counts"][table]
        assert all(a <= b for a,b in actual)
        assert all(a[1] < b[0] for a,b in zip(actual, actual[1:]))


def verify_manifest():
    manifest = load(HERE / "artifact-manifest.json")
    excluded = {"artifact-manifest.json", "root-verification.json"}
    actual = {str(p.relative_to(HERE)): sha(p) for p in HERE.rglob("*")
        if p.is_file() and str(p.relative_to(HERE)) not in excluded}
    assert actual == manifest["sha256"]
    assert len(actual) == manifest["artifact_count"]
    return len(actual)

def verify_profile_rows(folder, kind):
    rows = list(csv.DictReader((folder / "raw.csv").open()))
    commands = (folder / "commands.txt").read_text().splitlines()
    assert len(commands) == len(rows) > 0
    for row, command in zip(rows, commands):
        if kind == "abab":
            stem = folder / (row["side"] + "-r" + row["round"] + "-" + row["case"])
        elif kind == "evaluation":
            stem = folder / (row["phase"] + "-" + row["case"])
        else:
            stem = folder / ("parse-" + row["case"])
        lines = stem.with_suffix(".stdout").read_text().splitlines()
        assert len(lines) == 2 and lines[0].startswith("config ") and lines[1].startswith("result ")
        config, result = [dict(token.split("=", 1) for token in line.split()[1:]) for line in lines]
        for key, value in (config | result).items():
            assert row[key] == value, (folder, row["case"], key)
        assert row["status"] == stem.with_suffix(".status").read_text().strip() == "0"
        assert not stem.with_suffix(".stderr").read_text()
        timing = stem.with_suffix(".time").read_text()
        assert re.search(r"Exit status:\s*0\s*$", timing)
        assert row["max_rss_kib"] == re.search(r"Maximum resident set size \(kbytes\):\s*(\d+)", timing)[1]
        assert row["warmups"] == "3" and row["iterations"] == "15"
        expected = int(row["repeat"]) if row["expected_success"] == "true" else 0
        assert int(row["successes_p50"]) == int(row["successes_max"]) == expected
        if kind == "evaluation":
            assert int(row["refusals_p50"]) == int(row["refusals_max"]) == int(row["repeat"]) - expected
        assert row["live_before_p50"] == row["live_after_p50"] == row["live_after_max"]
        assert row["checksum_p50"] == row["checksum_max"]
        assert 0 < int(row["p50_ns"]) <= int(row["p95_ns"]) <= int(row["p99_ns"])
        argv = shlex.split(command)
        assert argv[:3] == ["taskset", "-c", "2"]
        for key in ["case", "repeat", "warmups", "iterations"]:
            assert argv[argv.index("--" + key) + 1] == row[key]
        if kind == "evaluation":
            assert argv[argv.index("--phase") + 1] == row["phase"]
    return rows


def verify_profiles(receipt):
    performance = HERE / "performance"
    binary_checks = load(HERE / "gates/root-binary-verification.json")["binaries"]
    base = load(HERE / "specification.json")["base_commit"]
    summaries = {}
    all_rows = {}
    for variant in ["baseline", "ascii-candidate", "ast-candidate", "evaluation-candidate"]:
        folder = performance / variant
        evaluation = variant == "evaluation-candidate"
        harness = performance / ("evaluation-harness" if evaluation else "ast-harness")
        runner = runpy.run_path(str(harness / "run.py"))
        cases = runner["CASES"]
        phases = runner["PHASES"] if evaluation else ["parse"]
        rows = verify_profile_rows(folder, "evaluation" if evaluation else "ast")
        expected = {(phase, case, str(repeat)) for phase in phases for case, repeat in cases}
        assert {(r.get("phase", "parse"), r["case"], r["repeat"]) for r in rows} == expected
        assert len(rows) == len(expected)
        group = load(folder / "group.json")
        assert group["case_count"] == len(cases)
        assert group["cases"] == [{"case": case, "repeat": repeat} for case, repeat in cases]
        binary = load(folder / "binary-provenance.json")
        assert binary["sha256"] == binary_checks[variant]["sha256"]
        assert binary["size_bytes"] == binary_checks[variant]["bytes"]
        assert binary["build_status"] == 0 and binary["base_commit"] == base
        for line in (folder / "harness-sha256.txt").read_text().splitlines():
            digest, name = line.split(maxsplit=1)
            assert sha(harness / name) == digest, (variant, name)
        build = load(folder / "build-command.json")
        assert build["status"] == 0
        assert build["command"] == ["cargo", "build", "--locked", "--offline", "--release", "--manifest-path", str(harness.relative_to(ROOT) / "Cargo.toml")]
        sources = load(folder / "source-sha256.json")
        if variant in ["ast-candidate", "evaluation-candidate"]:
            assert sources["source_before"] == sources["source_after"] == receipt["source_after"]
        elif variant == "baseline":
            assert sources["source_before"] == sources["source_after"]
            for name, digest in sources["source_after"].items():
                content = subprocess.check_output(["git", "show", base + ":" + name], cwd=ROOT)
                assert hashlib.sha256(content).hexdigest() == digest
        else:
            baseline = load(performance / "baseline/source-sha256.json")["source_after"]
            assert sources["source_before"] == baseline
            assert set(sources["source_after"]) == set(baseline)
            changed = [name for name, digest in sources["source_after"].items() if digest != baseline[name]]
            names = load(HERE / "names-verification.json")
            assert changed == [names["helper"]]
        summaries[variant] = len(rows)
        all_rows[variant] = {r["case"]: r for r in rows}
    # Parser semantics and allocations must match in every comparison lane.
    for variant in ["ascii-candidate", "ast-candidate"]:
        for case, row in all_rows[variant].items():
            old = all_rows["baseline"][case]
            for key in ["checksum_p50", "successes_p50", "alloc_calls_p50", "requested_bytes_p50", "peak_live_delta_p50"]:
                assert row[key] == old[key], (variant, case, key)
    ascii_rows = verify_profile_rows(performance / "ascii-candidate-abab", "abab")
    assert len(ascii_rows) == 24
    summaries["ascii-candidate-abab"] = len(ascii_rows)
    # Final regression repeats have a reduced CSV schema; check the full stdout too.
    folder = performance / "ast-abab"
    rows = list(csv.DictReader((folder / "raw.csv").open()))
    sequence = load(folder / "sequence.json")["rows"]
    assert len(rows) == len(sequence) == 16
    assert {(r["round"], r["variant"], r["case"]) for r in rows} == {
        (str(round), variant, case) for round in range(1, 5)
        for variant in ["baseline", "candidate"] for case in ["expr-array-1024", "expr-reference-1024"]}
    for row, record in zip(rows, sequence):
        stem = folder / (row["variant"] + "-r" + row["round"] + "-" + row["case"])
        lines = stem.with_suffix(".stdout").read_text().splitlines()
        assert len(lines) == 2
        values = {}
        for line in lines:
            values.update(dict(token.split("=", 1) for token in line.split()[1:]))
        for key in row.keys() & values.keys():
            assert row[key] == values[key], (stem, key)
        assert values["successes_p50"] == values["successes_max"] == "32"
        assert values["live_before_p50"] == values["live_after_p50"] == values["live_after_max"]
        assert values["checksum_p50"] == values["checksum_max"] == all_rows["baseline"][row["case"]]["checksum_p50"]
        assert row["status"] == stem.with_suffix(".status").read_text().strip() == "0"
        assert record["status"] == 0 and record["case"] == row["case"] and str(record["round"]) == row["round"]
        assert record["variant"] == row["variant"]
        assert not stem.with_suffix(".stderr").read_text()
        timing = stem.with_suffix(".time").read_text()
        assert re.search(r"Exit status:\s*0\s*$", timing)
        assert row["max_rss_kib"] == re.search(r"Maximum resident set size \(kbytes\):\s*(\d+)", timing)[1]
        argv = record["command"]
        assert argv[:3] == ["taskset", "-c", "2"]
        for key in ["case", "repeat", "warmups", "iterations"]:
            assert argv[argv.index("--" + key) + 1] == row[key]
    summaries["ast-abab"] = len(rows)
    return summaries


def verify_counters():
    folder = HERE / "performance/perf-stat"
    binary = load(HERE / "gates/root-binary-verification.json")["binaries"]["evaluation-candidate"]
    cases = set()
    for metadata in folder.glob("*.metadata.json"):
        record = load(metadata)
        stem = metadata.name.removesuffix(".metadata.json")
        assert record["status"] == 0
        assert (folder / (stem + ".status")).read_text().strip() == "0"
        assert record["binary_sha256"] == binary["sha256"]
        assert record["binary_size_bytes"] == binary["bytes"]
        for stream in ["stdout", "stderr"]:
            assert sha(folder / (stem + "." + stream)) == record[stream + "_sha256"]
        argv = record["command"]
        assert argv[:9] == ["perf", "stat", "-x,", "-e", "cycles,instructions,branches,branch-misses", "--", "taskset", "-c", "2"]
        assert argv[9] == record["binary"]
        assert argv[argv.index("--case") + 1] == record["case"]
        assert argv[argv.index("--phase") + 1] == "evaluate"
        cases.add(record["case"])
        lines = (folder / (stem + ".stdout")).read_text().splitlines()
        assert len(lines) == 2
        values = {}
        for line in lines:
            values.update(dict(token.split("=", 1) for token in line.split()[1:]))
        assert values["case"] == record["case"]
        assert values["live_before_p50"] == values["live_after_p50"] == values["live_after_max"]
        assert values["successes_p50"] == values["successes_max"] == "2"
        counters = list(csv.reader((folder / (stem + ".stderr")).read_text().splitlines()))
        assert {r[2] for r in counters if r} == {"cycles", "instructions", "branches", "branch-misses"}
        assert all(int(r[0]) > 0 for r in counters if r)
    assert cases == {"eval-flat-4096", "eval-utf8-number-left-4096", "eval-coerce-4096"}
    return len(cases)


def main():
    receipt, requirements = verify_gates()
    verify_sources(receipt)
    verify_specification()
    profiles = verify_profiles(receipt)
    counters = verify_counters()
    count = verify_manifest()
    result = {"status": 0, "tests": requirements["tests"], "targets": requirements["targets"],
              "doctests": requirements["doctests"], "profile_rows": profiles,
              "source_files": len(receipt["source_after"]), "artifacts": count,
              "counter_captures": counters}
    (HERE / "root-verification.json").write_text(json.dumps(result, indent=2) + "\n")
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
