#!/usr/bin/env python3
"""Verify retained roman evaluator source, gates, profiles and cleanup."""
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
    assert changed == set(replay["changed_files"]) == {
        "crates/litchi-ods/docs/FEATURE_MATRIX.md",
        "crates/litchi-ods/src/codec/formula/evaluation.rs",
        "crates/litchi-ods/src/codec/formula/evaluation/roman.rs",
        "crates/litchi-ods/tests/ods_formula_roman_evaluation.rs",
    }


def verify_gates():
    receipt = load(HERE / "gates/results.json")
    assert receipt["sources_unchanged"]
    assert receipt["source_before"] == receipt["source_after"]
    essential = {"crates/litchi-ods/src/codec/formula/evaluation.rs", "crates/litchi-ods/tests/ods_formula_evaluation.rs",
                 "crates/litchi-ods/src/codec/formula.rs", "crates/litchi-ods/src/codec/formula/expression.rs",
                 "crates/litchi-ods/src/codec/formula/expression/names.rs", "crates/litchi-ods/src/codec/formula/reference.rs",
                 "crates/litchi-ods/src/codec/formula/reference/iri.rs", "crates/litchi-ods/tests/ods_formula_expressions.rs"}
    essential.update({"crates/litchi-ods/src/codec/formula/evaluation/roman.rs",
                      "crates/litchi-ods/tests/ods_formula_roman_evaluation.rs",
                      "crates/litchi-ods/src/codec/formula/evaluation/radix.rs",
                      "crates/litchi-ods/tests/ods_formula_radix_evaluation.rs"})
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
    assert sha(ROOT / spec["source"]) == spec["source_sha256"]
    with zipfile.ZipFile(ROOT / spec["source"]) as archive:
        assert hashlib.sha256(archive.read(spec["entry"])).hexdigest() == spec["entry_sha256"]


def verify_reviews(receipt):
    for report in ['spec-review.md', 'implementation-review.md']:
        text = (HERE / report).read_text()
        for source in ['crates/litchi-ods/src/codec/formula/evaluation.rs',
                       'crates/litchi-ods/src/codec/formula/evaluation/roman.rs',
                       'crates/litchi-ods/tests/ods_formula_roman_evaluation.rs']:
            assert receipt['source_after'][source] in text, (report, source)
    numerical = load(HERE / 'numerical-review/receipt.json')
    assert numerical['execution']['status'] == 0
    assert numerical['script_sha256'] == sha(HERE / 'numerical-review/check_roman.py')
    for label, source in {
        'evaluation_dispatch': 'crates/litchi-ods/src/codec/formula/evaluation.rs',
        'roman_module': 'crates/litchi-ods/src/codec/formula/evaluation/roman.rs',
        'roman_tests': 'crates/litchi-ods/tests/ods_formula_roman_evaluation.rs',
    }.items():
        assert numerical['source_sha256'][label] == receipt['source_after'][source]
    spec = load(HERE / 'specification.json')
    assert numerical['source_sha256']['specification'] == sha(HERE / 'specification.json')
    assert numerical['source_sha256']['normative_archive'] == spec['source_sha256']
    assert numerical['source_sha256']['normative_entry'] == spec['entry_sha256']
    assert numerical['scope'] == {'min': 0, 'max': 3999, 'count': 4000}
    assert numerical['source_like_simplified_checks'] == {
        'checked_numbers': 4000, 'roundtrip_and_unrestricted_distance': True}
    for kind in ['canonical_signed_bfs', 'unrestricted_signed_bfs']:
        assert numerical[kind]['reached_scope']
        assert numerical[kind]['max_scope_distance'] == 11
    assert numerical['canonical_signed_bfs']['distance_matches_unrestricted']
    assert numerical['canonical_signed_bfs']['roundtrip_witnesses']
    assert set(numerical['source_like_greedy_checks']) == {'0', '1', '2', '3'}
    assert all(item['roundtrip_and_shape'] for item in numerical['source_like_greedy_checks'].values())
    return numerical['scope']['count']


def verify_manifest():
    manifest = load(HERE / "artifact-manifest.json")
    excluded = {"artifact-manifest.json", "root-verification.json"}
    actual = {str(p.relative_to(HERE)): sha(p) for p in HERE.rglob("*")
              if p.is_file() and str(p.relative_to(HERE)) not in excluded}
    assert {"README.md", "spec-review.md", "implementation-review.md", "performance/report.md", "performance/harness-review.md", "performance/environment.json"} <= set(actual)
    assert actual == manifest["sha256"]
    assert len(actual) == manifest["artifact_count"]
    return len(actual)


def verify_profile_rows(folder):
    rows = list(csv.DictReader((folder / "raw.csv").open()))
    commands = (folder / "commands.txt").read_text().splitlines()
    assert len(commands) == len(rows) > 0
    for row, command in zip(rows, commands):
        stem = folder / (row["revision"] + "-" + row["phase"] + "-" + row["case"])
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
        assert int(row["refusals_p50"]) == int(row["refusals_max"]) == int(row["repeat"]) - expected
        assert row["live_before_p50"] == row["live_after_p50"] == row["live_after_max"]
        assert row["checksum_p50"] == row["checksum_max"]
        assert 0 < int(row["p50_ns"]) <= int(row["p95_ns"]) <= int(row["p99_ns"])
        argv = shlex.split(command)
        assert argv[:3] == ["taskset", "-c", "6"]
        for key in ["case", "repeat", "warmups", "iterations"]:
            assert argv[argv.index("--" + key) + 1] == row[key]
        for key in ["revision", "phase", "group"]:
            assert argv[argv.index("--" + key) + 1] == row[key]
    return rows


def verify_profiles(receipt):
    performance = HERE / "performance"
    harness = performance / "roman-harness"
    previous_harness = performance / 'exploration/roman-harness'
    affinity = load(performance / 'final-harness-affinity.json')
    assert affinity['status'] == 0 and affinity['final_cpu'] == 6 and affinity['old_cpu'] == 2
    assert affinity['final_runner_sha256'] == sha(harness / 'run.py')
    assert affinity['old_runner_sha256'] == sha(previous_harness / 'run.py')
    assert (previous_harness / 'run.py').read_bytes().replace(b'        "2",', b'        "6",') == (harness / 'run.py').read_bytes()
    for name in ['Cargo.toml', 'Cargo.lock', 'src/main.rs']:
        assert (harness / name).read_bytes() == (previous_harness / name).read_bytes()
    assert affinity['rust_harness_sha256'] == sha(harness / 'src/main.rs')
    runner = runpy.run_path(str(harness / "run.py"))
    binaries = load(HERE / "gates/root-binary-verification.json")["binaries"]
    base = load(HERE / "specification.json")["base_commit"]
    rows_by_group = {}
    for variant in ["baseline", "candidate"]:
        folder = performance / variant
        binary = load(folder / "binary-provenance.json")
        assert binary["sha256"] == binaries[variant]["sha256"]
        assert binary["size_bytes"] == binaries[variant]["bytes"]
        assert binary["build_status"] == 0
        assert base.startswith(binary["base_commit"])
        for line in (folder / "harness-sha256.txt").read_text().splitlines():
            digest, name = line.split(maxsplit=1)
            assert sha(harness / name) == digest
        source = load(folder / "source-sha256.json")
        assert source["source_before"] == source["source_after"]
        assert source["sources_unchanged"] and source["status"] == 0
        if variant == "candidate":
            assert source["source_after"] == receipt["source_after"]
        else:
            for name, digest in source["source_after"].items():
                if name == "Cargo.lock":
                    # The workspace lock is a pre-existing ignored input;
                    # retain it explicitly instead of claiming it is in Git.
                    assert sha(HERE / "gates/workspace-Cargo.lock") == digest
                    assert digest == receipt["source_after"][name]
                    continue
                content = subprocess.check_output(["git", "show", base + ":" + name], cwd=ROOT)
                assert hashlib.sha256(content).hexdigest() == digest
        for group in (["comparable"] if variant == "baseline" else ["comparable", "roman"]):
            rows = verify_profile_rows(folder / group)
            cases = runner["COMPARABLE_CASES" if group == "comparable" else "ROMAN_CASES"]
            expected = {(variant, group, phase, case, str(repeat)) for phase in runner["PHASES"] for case, repeat in cases}
            assert len(rows) == len(expected)
            assert {(r["revision"], r["group"], r["phase"], r["case"], r["repeat"]) for r in rows} == expected
            rows_by_group[variant + "/" + group] = rows
    old = {(r["phase"], r["case"]): r for r in rows_by_group["baseline/comparable"]}
    for row in rows_by_group["candidate/comparable"]:
        previous = old[(row["phase"], row["case"])]
        for key in ["input_bytes", "expected_success", "successes_p50", "refusals_p50", "failure", "checksum_p50", "output_reserved_bytes_p50", "alloc_calls_p50", "requested_bytes_p50", "peak_live_delta_p50"]:
            assert row[key] == previous[key], (row["phase"], row["case"], key)
    return {group: len(rows) for group, rows in rows_by_group.items()}


def verify_sequence(folder, rows):
    sequence = load(folder / 'sequence.json')
    assert len(sequence) == len(rows)
    previous_finish = ''
    for row, item in zip(rows, sequence):
        for key in ['revision', 'group', 'phase', 'case', 'repeat', 'status']:
            assert str(item[key]) == row[key]
        assert previous_finish <= item['started_at'] <= item['finished_at']
        previous_finish = item['finished_at']
        stem = '-'.join(row[k] for k in ['revision', 'phase', 'case'])
        for suffix in ['stdout', 'stderr', 'time']:
            assert item[suffix] == stem + '.' + suffix


def verify_abab(parent):
    import math
    import statistics
    flags = load(parent / 'initial-flags.json')
    assert len({(x['phase'], x['case']) for x in flags}) == len(flags)
    combined = []
    for number in range(1, 5):
        folder = parent / 'abab' / f'r{number}'
        rows = verify_profile_rows(folder)
        verify_sequence(folder, rows)
        expected = [(v, f['phase'], f['case'], str(f['repeat']))
                    for f in flags for v in ['baseline', 'candidate']]
        assert [(r['revision'], r['phase'], r['case'], r['repeat']) for r in rows] == expected
        for a, b in zip(rows[::2], rows[1::2]):
            for key in ['checksum_p50', 'failure', 'expected_success', 'alloc_calls_p50',
                        'requested_bytes_p50', 'peak_live_delta_p50', 'output_reserved_bytes_p50']:
                assert a[key] == b[key], (parent, a['case'], key)
        combined.extend(dict(round=str(number), **r) for r in rows)
    assert list(csv.DictReader((parent / 'abab/raw.csv').open())) == combined
    summary = load(parent / 'abab/summary.json')
    assert len(summary) == len(flags)
    for flag, item in zip(flags, summary):
        assert (item['phase'], item['case'], item['repeat']) == (flag['phase'], flag['case'], int(flag['repeat']))
        selected = [r for r in combined if r['phase'] == flag['phase'] and r['case'] == flag['case']]
        for metric in ['p50_ns', 'p95_ns', 'p99_ns', 'max_rss_kib', 'alloc_calls_p50',
                       'requested_bytes_p50', 'peak_live_delta_p50']:
            med = {v: statistics.median(float(r[metric]) for r in selected if r['revision'] == v)
                   for v in ['baseline', 'candidate']}
            delta = (med['candidate'] / med['baseline'] - 1) * 100 if med['baseline'] else 0
            assert item[metric]['baseline'] == med['baseline']
            assert item[metric]['candidate'] == med['candidate']
            if med['baseline'] == 0:
                assert med['candidate'] == 0 and item[metric]['delta_pct'] is None
            else:
                assert math.isclose(item[metric]['delta_pct'], delta, abs_tol=1e-10)
    old = {(r['phase'], r['case']): r for r in verify_profile_rows(parent / 'baseline/comparable')}
    new = {(r['phase'], r['case']): r for r in verify_profile_rows(parent / 'candidate/comparable')}
    metrics = ['p50_ns', 'p95_ns', 'p99_ns', 'max_rss_kib']
    deltas = {key: {m: (float(new[key][m]) / float(row[m]) - 1) * 100 if float(row[m]) else 0
                    for m in metrics} for key, row in old.items()}
    expected = {key for key, delta in deltas.items() if any(v > 5 for v in delta.values())}
    assert {(f['phase'], f['case']) for f in flags} == expected
    for f in flags:
        for metric in metrics:
            assert math.isclose(f['delta_pct'][metric], deltas[f['phase'], f['case']][metric], abs_tol=1e-10)
    return len(combined)


def verify_build(folder, variant):
    binary = load(folder / 'binary-provenance.json')
    custody = load(HERE / 'gates/root-binary-verification.json')['binaries'][variant]
    assert binary['sha256'] == custody['sha256'] and binary['size_bytes'] == custody['bytes']
    source = load(folder / 'source-sha256.json')
    assert source['source_before'] == source['source_after']
    assert source['sources_unchanged'] and source['status'] == binary['build_status'] == 0
    build = load(folder / 'build-command.json')
    assert build['status'] == 0
    assert build['command'] == source['command']
    assert '--locked' in build['command'] and '--offline' in build['command'] and '--release' in build['command']
    assert source['source_before_captured_at'] <= source['started_at'] <= source['finished_at'] <= source['source_after_captured_at']
    for line in (folder / 'harness-sha256.txt').read_text().splitlines():
        digest, name = line.split(maxsplit=1)
        assert sha(HERE / 'performance/roman-harness' / name) == digest
    custody = load(folder / 'harness-custody.json')
    assert custody['before'] == custody['after']
    assert custody['before'] == {f: sha(HERE / 'performance/roman-harness' / f) for f in ['Cargo.toml', 'Cargo.lock', 'run.py', 'src/main.rs']}
    for group in ['comparable', 'roman']:
        path = folder / group
        if not path.exists(): continue
        capture = load(path / 'capture.json')
        assert capture['status'] == 0 and capture['started_at'] <= capture['finished_at']
        assert capture['binary_sha256'] == binary['sha256']
        assert capture['raw_sha256'] == sha(path / 'raw.csv')
    return source


def verify_counters(folder=None, expected_count=5):
    binaries = load(HERE / 'gates/root-binary-verification.json')['binaries']
    folder = folder or HERE / 'performance/perf-stat'
    files = sorted(folder.glob('*.metadata.json'))
    assert len(files) == expected_count
    for path in files:
        receipt = load(path); stem = path.with_name(path.name.removesuffix('.metadata.json'))
        stdout = stem.with_suffix('.stdout'); stderr = stem.with_suffix('.stderr')
        assert receipt['status'] == 0 and stem.with_suffix('.status').read_text().strip() == '0'
        assert sha(stdout) == receipt['stdout_sha256'] and sha(stderr) == receipt['stderr_sha256']
        argv = receipt['argv']
        assert argv[:8] == ['perf', 'stat', '-x,', '-e', 'cycles,instructions,branches,branch-misses', '--', 'taskset', '-c']
        assert argv[8] == '6'
        variant = argv[argv.index('--revision') + 1]
        assert receipt['binary_sha256'] == binaries[variant]['sha256']
        assert receipt['binary_bytes'] == binaries[variant]['bytes']
        assert receipt['started_at'] <= receipt['finished_at']
        counters = {r[2]: float(r[0]) for r in csv.reader(stderr.read_text().splitlines()) if len(r) >= 3}
        assert set(counters) == {'cycles', 'instructions', 'branches', 'branch-misses'}
        assert all(v >= 0 for v in counters.values()) and counters['cycles'] > 0 and counters['instructions'] > 0
        lines = stdout.read_text().splitlines(); assert len(lines) == 2
        config, result = [dict(token.split('=', 1) for token in line.split()[1:]) for line in lines]
        for key in ['revision', 'group', 'phase', 'case', 'warmups', 'iterations', 'repeat']:
            assert config[key] == argv[argv.index('--' + key) + 1]
        assert config['warmups'] == '3' and config['iterations'] == '15'
        assert result['live_before_p50'] == result['live_after_p50'] == result['live_after_max']
        assert result['successes_p50'] == config['repeat'] and result['refusals_p50'] == '0'
    return len(files)



def verify_cleanup():
    cleanup = load(HERE / 'gates/cleanup.json')
    assert cleanup['all_removed_paths_absent'] and cleanup['remaining_process_references'] == []
    assert cleanup['reclaimed_allocated_bytes'] > 0
    recovery = load(HERE / 'gates/recovery-manifest.json')
    assert recovery['archive_sha256'] == cleanup['unique_archive_sha256']
    assert len(recovery['files']) == cleanup['unique_file_count']
    assert set(cleanup['removed_directories']) == {'/home/zhuhe/code/litchi-roman-worktree', '/home/zhuhe/code/litchi-roman-target', '/home/zhuhe/code/litchi-roman-tmp'}
    assert all(not Path(p).exists() for p in cleanup['removed_directories'])
    binaries = load(HERE / 'gates/root-binary-verification.json')
    assert binaries['removed_after_verification_at']
    assert len(binaries['removed_binaries']) == len(binaries['binaries']) == 2
    for item in binaries['removed_binaries']:
        original = binaries['binaries'][item['variant']]
        for key in ['path', 'sha256', 'bytes']: assert item[key] == original[key]
        assert not (ROOT / item['path']).exists()
    return cleanup['reclaimed_allocated_bytes']


def main():
    import datetime
    receipt, requirements = verify_gates()
    verify_sources(receipt); verify_specification()
    numerical_values = verify_reviews(receipt)
    profiles = verify_profiles(receipt)
    for variant in ['baseline', 'candidate']:
        verify_build(HERE / 'performance' / variant, variant)
    abab = verify_abab(HERE / 'performance')
    counters = verify_counters()
    investigation_counters = verify_counters(HERE / 'performance/investigation-counters', 4)
    reclaimed = verify_cleanup()
    artifacts = verify_manifest()
    result = {'status': 0, 'verified_at': datetime.datetime.now(datetime.timezone.utc).isoformat(),
              'tests': requirements['tests'], 'targets': requirements['targets'], 'doctests': requirements['doctests'],
              'source_inputs': len(receipt['source_after']), 'profiles': profiles,
              'review_source_hashes_checked': True,
              'independent_python_algorithm_values': numerical_values,
              'final_abab_rows': abab,
              'counter_captures': counters, 'reclaimed_allocated_bytes': reclaimed, 'artifacts': artifacts}
    result['investigation_counter_captures'] = investigation_counters
    (HERE / 'root-verification.json').write_text(json.dumps(result, indent=2) + '\n')
    print(json.dumps(result, indent=2))


if __name__ == '__main__': main()
