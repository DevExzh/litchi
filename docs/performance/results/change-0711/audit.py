#!/usr/bin/env python3
"""Audit the rejected pilot and replay its analyzer after owned cleanup."""
import hashlib
import importlib.util
import json
from pathlib import Path
import subprocess
import tempfile

P = Path(__file__).resolve().parent
ROOT = P.parents[3]

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def read(path):
    return json.loads(path.read_text())

def main():
    spec = importlib.util.spec_from_file_location('custody0711', P/'custody.py')
    custody = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(custody)
    baseline = read(P/'source-baseline.json')
    candidate = read(P/'source-candidate.json')
    assert custody.census() == baseline == read(P/'source-final.json')
    target = 'crates/litchi-docx/src/alt/codec.rs'
    assert {n for n in set(baseline)|set(candidate) if baseline.get(n) != candidate.get(n)} == {target}
    assert sha(P/'baseline-codec.rs') == baseline[target]
    assert sha(P/'candidate-codec.rs') == candidate[target]
    for name, digest in read(P/'constraints.json').items():
        assert sha(ROOT/name) == digest, name
    for filename, folder in [('capture-freeze.json', P), ('oracle-freeze.json', P/'oracle')]:
        for name, digest in read(P/filename).items():
            assert sha(folder/name) == digest, name
    cleanup = read(P/'cleanup.json')
    expected_paths = ['/home/zhuhe/code/litchi-target-0711', '/home/zhuhe/code/litchi-0711-bin', '/home/zhuhe/code/litchi-0711-fs']
    assert cleanup['owned_paths'] == expected_paths
    assert cleanup['owned_paths_absent'] and all(not Path(n).exists() for n in expected_paths)
    assert len(cleanup['binaries']) == 6
    assert len({r['path'] for r in cleanup['binaries']}) == 6
    for stage in ['baseline', 'candidate']:
        for build in read(P/f'build-{stage}.json'):
            assert build['exit_code'] == 0
            assert dict(path=build['binary'], sha256=build['binary_sha256'], bytes=build['binary_bytes']) in cleanup['binaries']
        folder = P/'oracle'/stage
        result = read(folder/'result.json')
        build = read(folder/'build.json')
        assert result['exit_code'] == build['exit_code'] == 0
        assert result['binary'] in cleanup['binaries']
        assert result['source_sha256'] == build['source_sha256'] == sha(folder/'source.json') == sha(P/f'source-{stage}.json')
        assert result['probe_sha256'] == build['probe_sha256'] == sha(folder/'probe.json')
        for name, digest in read(folder/'probe.json').items():
            assert sha(P/'oracle'/name) == digest
        assert sha(folder/'build.log') == build['log_sha256']
        for name, digest in result['artifacts'].items():
            assert sha(folder/name) == digest
        report = read(folder/'report.json')
        assert report['case_count'] == len(report['cases']) == 49
        assert sum(c['outcome']['kind'] == 'success' for c in report['cases']) == 28
        assert sum(c['outcome']['kind'] == 'error' for c in report['cases']) == 21
    assert (P/'oracle/baseline/report.json').read_bytes() == (P/'oracle/candidate/report.json').read_bytes()
    analysis = read(P/'analysis.json')
    assert analysis['verification']['child_receipts_verified'] == 32
    decision = analysis['decision']
    assert not decision['accepted'] and decision['deterministic_output_parity_pass']
    failed = [r for r in decision['hard_gates'] if not r['pass']]
    assert len(decision['hard_gates']) == 48 and len(failed) == 1
    assert failed[0]['name'] == 'pair-2/numbered-list/edit/p50/improvement'
    assert 2.88 < failed[0]['observed_improvement_percent'] < 2.89
    assert read(P/'pilot-gate.json')['analysis_sha256'] == sha(P/'analysis.json')
    assert read(P/'pilot-gate.json')['status'] == read(P/'disposition.json')['decision'] == 'rejected'
    negative = read(P/'negative-checks.json')
    assert negative['status'] == 'pass' and negative['retained_inputs_unchanged']
    assert negative['analyzer_sha256'] == sha(P/'analyze.py')
    assert negative['analysis_sha256'] == sha(P/'analysis.json')
    assert len(negative['checks']) == 4 and all(r['rejected'] for r in negative['checks'])
    evidence = read(P/'evidence/results.json')
    assert len(evidence) == 6
    for row in evidence:
        assert row['exit_code'] == 0
        assert row['source_manifest_sha256'] == sha(P/'source-final.json')
        assert row['log_sha256'] == sha(P/'evidence'/(row['name']+'.log'))
    final = read(P/'final-report-gate.json')
    assert final['exit_code'] == 0 and final['log_sha256'] == sha(P/'final-report-gate.log')
    for name, digest in final['docs'].items():
        assert sha(ROOT/name) == digest
    with tempfile.TemporaryDirectory(prefix='litchi-0711-audit-') as directory:
        output = Path(directory)/'analysis.json'
        subprocess.run(['python3', '-B', str(P/'analyze.py'), '--output', str(output)], cwd=ROOT, check=True, stdout=subprocess.DEVNULL)
        assert output.read_bytes() == (P/'analysis.json').read_bytes()
    assert not (P/'quality.json').exists()
    assert not list(P.glob('*.callgrind*'))
    print('PASS: 32 children, 49-case parity, rejection, restored source, six gates, final docs, cleanup and exact replay')

if __name__ == '__main__':
    main()
