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
assert len(results) == 48 and sum(int(x[0]) for x in results) == 815
assert all(x[1:] == ('0', '0') for x in results)
red = json.loads((HERE / 'gates/allocator-baseline-reproduction.json').read_text())
assert red['expected_status'] == red['observed_status'] == 101
assert red['allocator_failure_observed'] and red['source_hashes_unchanged']
assert red['formula_sha256'] == hashlib.sha256(subprocess.check_output(['git', 'show', gates['head'] + ':crates/litchi-ods/src/codec/formula.rs'], cwd=ROOT)).hexdigest()
assert 'swallowed allocation failure' in red['output']
perf = HERE / 'performance'
for name, expected in json.loads((perf / 'baseline/source-sha256.json').read_text()).items():
    content = subprocess.check_output(['git', 'show', gates['head'] + ':' + name], cwd=ROOT)
    assert hashlib.sha256(content).hexdigest() == expected, name
candidate_sources = json.loads((perf / 'candidate/source-sha256.json').read_text())
assert candidate_sources['base_commit'] == gates['head']
assert candidate_sources['source_after'] == gates['source_after']
for directory in ['baseline', 'candidate']:
    for line in (perf / directory / 'harness-sha256.txt').read_text().splitlines():
        expected, name = line.split('  ', 1)
        assert sha(ROOT / name) == expected, name

all_rows = {}
for directory, count in [('baseline', 12), ('candidate', 12), ('baseline-coverage', 16), ('candidate-coverage', 16)]:
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
for left, right in [('baseline', 'candidate'), ('baseline-coverage', 'candidate-coverage')]:
    for a, b in zip(all_rows[left], all_rows[right]):
        for key in ['case', 'input_bytes', 'repeat', 'expected_success']:
            assert a[key] == b[key], key
        if a['expected_success'] == 'true':
            assert a['checksum_p50'] == b['checksum_p50']
isolated = json.loads((perf / 'candidate/isolated-checks.json').read_text())
assert len(isolated['checks']) == 4
counts = []
for check in isolated['checks']:
    assert check['status'] == 0
    matches = re.findall(r'test result: ok\. (\d+) passed; (\d+) failed; (\d+) ignored', (perf / check['log']).read_text())
    assert len(matches) == 1 and matches[0][1:] == ('0', '0')
    counts.append(int(matches[0][0]))
assert sorted(counts) == [4, 7, 10, 48]
binaries = json.loads((HERE / 'gates/root-binary-verification.json').read_text())['binaries']
for kind, receipt in binaries.items():
    assert (perf / kind / 'binary-sha256.txt').read_text().split()[0] == receipt['sha256']
    assert (perf / kind / 'build.status').read_text().strip() == '0'
sequence = json.loads((perf / 'paired/sequence.json').read_text())
assert len(sequence['jobs']) == 4
for job in sequence['jobs']:
    kind = job['label'].split('-')[0]
    folder = kind + ('-coverage' if job['group'] == 'coverage' else '')
    assert job['status'] == 0 and job['binary_sha256'] == binaries[kind]['sha256']
    assert job['raw_sha256'] == sha(perf / folder / 'raw.csv')
for kind in ['baseline', 'candidate']:
    for case in ['sum', 'local-refs-256']:
        stem = perf / 'paired' / ('perf-stat-' + kind + '-' + case)
        assert stem.with_suffix('.status').read_text().strip() == '0'
        counters = list(csv.reader(stem.with_suffix('.stderr').read_text().splitlines()))
        assert {r[2] for r in counters} == {'cycles', 'instructions', 'branches', 'branch-misses', 'cache-misses'}
        assert all(float(r[0]) > 0 for r in counters)
replay = json.loads((HERE / 'gates/root-patch-replay.json').read_text())
assert replay['base'] == gates['head'] and replay['verified']
assert replay['patch_sha256'] == sha(HERE / 'candidate.patch')
assert len(replay['source_after']) == 3
assert all(gates['source_after'][f] == digest for f, digest in replay['source_after'].items())
manifest_lines = (HERE / 'artifacts.sha256').read_text().splitlines()
assert len(manifest_lines) >= 240
for line in manifest_lines:
    expected, name = line.split('  ', 1)
    assert sha(HERE / name) == expected, name
print(json.dumps({'verified': True, 'tests': 815, 'targets': 48, 'performance_lanes': 56, 'allocation_failure_reproduced_and_fixed': True}))
