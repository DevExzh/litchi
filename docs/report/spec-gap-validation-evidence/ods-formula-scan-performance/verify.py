#!/usr/bin/env python3
"""Verify source-bound tokenizer gates and retained paired measurements."""
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

gates = json.loads((HERE / 'gates/results.json').read_text())
assert gates['sources_unchanged'] and gates['source_before'] == gates['source_after']
assert len(gates['commands']) == 5 and all(x['status'] == 0 for x in gates['commands'])
for name, expected in gates['source_after'].items():
    assert sha(ROOT / name) == expected, name
results = re.findall(r'test result: ok\. (\d+) passed; (\d+) failed; (\d+) ignored', (HERE / 'gates/test.log').read_text())
expected_tests = json.loads((HERE / 'requirements.json').read_text())
assert len(results) == expected_tests['targets'] and sum(int(x[0]) for x in results) == expected_tests['tests']
assert all(x[1:] == ('0', '0') for x in results)
perf = HERE / 'performance'
baseline_sources = json.loads((perf / 'baseline/source-sha256.json').read_text())
assert baseline_sources['commit'] == gates['head']
assert baseline_sources['source_before'] == baseline_sources['source_after']
for name, expected in baseline_sources['source_after'].items():
    if name.endswith('/ods_formula_scan_regression.rs'):
        assert expected == gates['source_after'][name]
    else:
        content = subprocess.check_output(['git', 'show', gates['head'] + ':' + name], cwd=ROOT)
        assert hashlib.sha256(content).hexdigest() == expected, name
baseline_check = json.loads((perf / 'baseline/isolated-scan-regression.json').read_text())
assert baseline_check['status'] == 0 and baseline_check['test_count'] == 6
assert 'test result: ok. 6 passed; 0 failed; 0 ignored' in (perf / 'baseline/isolated-scan-regression.log').read_text()
boundaries = json.loads((HERE / 'gates/boundaries.json').read_text())
assert boundaries['status'] == 0
assert boundaries['tool_sha256'] == sha(ROOT / 'tools/check_crate_boundaries.py')
assert 'crate boundaries valid for 65 workspace packages and 241 internal dependency declarations' in (HERE / 'gates/boundaries.log').read_text()
for kind in ['baseline', 'candidate', 'annotation-removals']:
    for line in (perf / kind / 'harness-sha256.txt').read_text().splitlines():
        expected, name = line.split('  ', 1)
        assert sha(ROOT / name) == expected, name
    assert (perf / kind / 'build.status').read_text().strip() == '0'
binaries = json.loads((HERE / 'gates/root-binary-verification.json').read_text())['binaries']
assert set(binaries) == {'baseline', 'candidate', 'annotation-removals'}
for kind, binary in binaries.items():
    assert binary['sha256'] == (perf / kind / 'binary-sha256.txt').read_text().split()[0]
candidate_sources = json.loads((perf / 'candidate/source-sha256.json').read_text())
assert candidate_sources['base_commit'] == gates['head']
assert candidate_sources['source_after'] == gates['source_after']
isolated = json.loads((perf / 'candidate/isolated-checks.json').read_text())
assert len(isolated['checks']) == 5 and isolated['expected_total'] == 76
counts = []
for check in isolated['checks']:
    assert check['status'] == 0
    matches = re.findall(r'test result: ok\. (\d+) passed; (\d+) failed; (\d+) ignored', (perf / check['log']).read_text())
    assert len(matches) == 1 and matches[0][1:] == ('0', '0')
    counts.append(int(matches[0][0]))
assert sorted(counts) == [4, 6, 7, 10, 49]
all_rows = {}
for directory, count in [('baseline', 12), ('candidate', 12), ('baseline-coverage', 16), ('candidate-coverage', 16), ('baseline-scaling', 8), ('candidate-scaling', 8)] + [('annotation-removals/' + name, 4) for name in ['round-1-baseline', 'round-2-annotation-removals', 'round-3-baseline', 'round-4-annotation-removals']] + [('local-refs-abab/' + name, 1) for name in ['round-1-baseline', 'round-2-candidate', 'round-3-baseline', 'round-4-candidate']]:
    folder = perf / directory
    rows = list(csv.DictReader((folder / 'raw.csv').open()))
    assert len(rows) == count
    all_rows[directory] = rows
    for row in rows:
        stem = folder / ('parse-' + row['case'])
        lines = stem.with_suffix('.stdout').read_text().splitlines()
        for prefix in ['config ', 'result ']:
            line = next(x for x in lines if x.startswith(prefix))
            parsed = dict(x.split('=', 1) for x in line.split()[1:])
            for key, value in parsed.items():
                if key in row:
                    assert row[key] == value, (directory, row['case'], key)
        time_text = stem.with_suffix('.time').read_text()
        rss = re.search(r'Maximum resident set size \(kbytes\):\s*(\d+)', time_text).group(1)
        assert rss == row['max_rss_kib']
        assert row['status'] == stem.with_suffix('.status').read_text().strip() == '0'
        assert stem.with_suffix('.stderr').read_text() == ''
        expected = int(row['repeat']) if row['expected_success'] == 'true' else 0
        assert int(row['successes_p50']) == int(row['successes_max']) == expected
        assert row['warmups'] == '3' and row['iterations'] == '15'
for left, right in [('baseline', 'candidate'), ('baseline-coverage', 'candidate-coverage'), ('baseline-scaling', 'candidate-scaling')]:
    for a, b in zip(all_rows[left], all_rows[right]):
        for key in ['case', 'input_bytes', 'repeat', 'expected_success']:
            assert a[key] == b[key], key
        if a['expected_success'] == 'true':
            assert a['checksum_p50'] == b['checksum_p50']
sequence = json.loads((perf / 'paired/sequence.json').read_text())
assert len(sequence['jobs']) == 6
for job in sequence['jobs']:
    kind = job['label'].split('-')[0]
    folder = kind + ('' if job['group'] == 'comparable' else '-' + job['group'])
    assert job['status'] == 0 and job['binary_sha256'] == binaries[kind]['sha256']
    assert job['raw_sha256'] == sha(perf / folder / 'raw.csv')
for directory in ['annotation-removals', 'local-refs-abab']:
    order = json.loads((perf / directory / 'sequence.json').read_text())['order']
    assert len(order) == 4
    for run in order:
        kind = 'baseline' if run['label'].endswith('baseline') else ('candidate' if directory == 'local-refs-abab' else directory)
        assert str(run['status']) == '0' and run['binary_sha256'] == binaries[kind]['sha256']
        if 'raw_sha256' in run:
            assert run['raw_sha256'] == sha(perf / directory / run['label'] / 'raw.csv')
        if directory == 'local-refs-abab':
            row = all_rows[directory + '/' + run['label']][0]
            for key, value in row.items():
                if key in run:
                    assert str(run[key]) == value, (directory, run['label'], key)
replay = json.loads((HERE / 'gates/root-patch-replay.json').read_text())
assert replay['verified'] and replay['base'] == gates['head']
assert replay['patch_sha256'] == sha(HERE / 'candidate.patch')
assert len(replay['source_after']) == 2
assert all(gates['source_after'][f] == digest for f, digest in replay['source_after'].items())
annotation = json.loads((HERE / 'gates/root-annotation-replay.json').read_text())
assert annotation['verified'] and annotation['base'] == gates['head']
assert annotation['patch_sha256'] == sha(perf / 'annotation-removals/annotation-removal.patch')
assert annotation['source_reference_sha256'] == json.loads((perf / 'annotation-removals/binary-provenance.json').read_text())['source_reference_sha256']
counters_dir = perf / 'paired/perf-stat'
metrics = json.loads((counters_dir / 'sequence.json').read_text())
assert len(metrics['jobs']) == 8
for job in metrics['jobs']:
    assert job['status'] == 0 and job['binary_sha256'] == binaries[job['side']]['sha256']
    assert (counters_dir / (job['label'] + '.status')).read_text().strip() == '0'
    for stream in ['stdout', 'stderr']:
        assert sha(counters_dir / job[stream]) == job[stream + '_sha256']
    rows = list(csv.reader((counters_dir / job['stderr']).read_text().splitlines()))
    assert {r[2] for r in rows} == {'cycles', 'instructions', 'branches', 'branch-misses', 'cache-misses'}
    assert all(float(r[0]) > 0 for r in rows)
manifest_lines = (HERE / 'artifacts.sha256').read_text().splitlines()
assert len(manifest_lines) >= 300
for line in manifest_lines:
    expected, name = line.split('  ', 1)
    assert sha(HERE / name) == expected, name
print(json.dumps({'verified': True, 'tests': expected_tests['tests'], 'targets': expected_tests['targets'], 'performance_lanes': 72, 'annotation_experiment_lanes': 16, 'followup_lanes': 4}))
