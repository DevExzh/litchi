#!/usr/bin/env python3
"""Verify retained gate hashes and every performance CSV field against raw output.

Run from the repository root. No build or timing run is performed.
"""
import csv
import hashlib
import json
import re
import subprocess
from pathlib import Path

EVIDENCE = Path(__file__).resolve().parent
ROOT = EVIDENCE.parents[3]


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def key_values(line):
    return dict(token.split('=', 1) for token in line.split()[1:])


def main():
    before = json.loads((EVIDENCE / 'gates/source-before.json').read_text())
    after = json.loads((EVIDENCE / 'gates/source-after.json').read_text())
    assert before == after
    for path, expected in after.items():
        assert sha(ROOT / path) == expected, path
    gates = json.loads((EVIDENCE / 'gates/gate-results.json').read_text())
    assert len(gates) == 5 and all(g['exit_code'] == 0 for g in gates)
    results = re.findall(r'test result: ok\. (\d+) passed',
                         (EVIDENCE / 'gates/all-targets.log').read_text())
    assert sum(map(int, results)) == 783 and len(results) == 46
    for variant in ('baseline', 'candidate'):
        directory = EVIDENCE / 'performance' / variant
        sources = json.loads((directory / 'source-before.json').read_text())
        assert sources == json.loads((directory / 'source-after.json').read_text())
        commit = (directory / 'commit.txt').read_text().strip()
        for path, expected in sources.items():
            if variant == 'baseline' and not path.startswith('docs/'):
                content = subprocess.check_output(
                    ['git', 'show', commit + ':' + path], cwd=ROOT)
                assert hashlib.sha256(content).hexdigest() == expected, path
            else:
                assert sha(ROOT / path) == expected, path
    replay = json.loads((EVIDENCE / 'performance/patch-replay.json').read_text())
    assert replay['source_hashes'] == after
    assert replay['patch_sha256'] == sha(EVIDENCE / 'performance/candidate.patch')
    isolated = EVIDENCE / 'performance/candidate'
    checks = json.loads((isolated / 'isolated-gates.json').read_text())
    assert len(checks['commands']) == 2
    assert all(check['exit_code'] == 0 for check in checks['commands'])
    for filename, count in (('formula-unit.log', 30), ('formula-integration.log', 7)):
        assert f'test result: ok. {count} passed' in (isolated / filename).read_text()
    counts = {}
    tables = {}
    for relative in ('baseline/raw.csv', 'candidate/raw.csv', 'candidate/new-names.csv'):
        path = EVIDENCE / 'performance' / relative
        rows = list(csv.DictReader(path.open()))
        tables[relative] = rows
        for row in rows:
            if 'name' in row:
                stem = row['workload'] + '-new-' + row['name'].replace('.', '_')
            else:
                stem = row['workload'] + '-' + row['case']
            raw = path.parent / stem
            lines = raw.with_suffix('.stdout').read_text().splitlines()
            config = key_values(next(x for x in lines if x.startswith('config ')))
            result = key_values(next(x for x in lines if x.startswith('result ')))
            measured = {**config, **result}
            measured['max_rss_kib'] = re.search(
                r'Maximum resident set size \(kbytes\):\s*(\d+)',
                raw.with_suffix('.time').read_text()).group(1)
            measured['status'] = raw.with_suffix('.status').read_text().strip()
            assert measured['status'] == '0'
            assert not raw.with_suffix('.stderr').read_text()
            assert config['workload'].lower() == row['workload']
            if 'case' in row:
                assert config['case'] == row['case']
            for key, value in row.items():
                if key not in ('case', 'name', 'workload'):
                    assert measured[key] == value, (relative, stem, key)
            if relative.startswith('candidate/') and row['workload'] == 'lookup':
                for key in ('alloc_calls_p50', 'alloc_calls_max',
                            'requested_bytes_p50', 'requested_bytes_max',
                            'peak_live_delta_p50', 'peak_live_delta_max'):
                    assert row[key] == '0', (stem, key)
            if 'name' in row:
                assert row['successes_p50'] == row['repeat']
                assert row['successes_max'] == row['repeat']
        counts[relative] = len(rows)
    assert counts == {'baseline/raw.csv': 23, 'candidate/raw.csv': 23,
                      'candidate/new-names.csv': 12}
    baseline = {r['case']: r for r in tables['baseline/raw.csv']}
    for candidate in tables['candidate/raw.csv']:
        previous = baseline[candidate['case']]
        for key in ('workload', 'input_bytes', 'repeat', 'warmups', 'iterations',
                    'successes_p50', 'successes_max', 'checksum_p50', 'checksum_max'):
            assert previous[key] == candidate[key], (candidate['case'], key)
    print(json.dumps({'source_hashes_match': True, 'gates_passed': 5,
                      'baseline_git_and_candidate_manifests_match': True,
                      'test_count': 783, 'test_targets': 46,
                      'isolated_formula_tests_passed': 37,
                      'csv_rows_verified_against_raw': counts,
                      'comparable_inputs_and_outputs_match': True,
                      'candidate_lookup_allocations_zero': True}, indent=2))


if __name__ == '__main__':
    main()
