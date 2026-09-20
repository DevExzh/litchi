#!/usr/bin/env python3
"""Audit the shared MCE pilot and exact evidence replay after owned cleanup."""
import difflib
import hashlib
import re
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
    spec = importlib.util.spec_from_file_location('custody0713', P/'custody.py')
    custody = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(custody)
    baseline = read(P/'source-baseline.json')
    candidate = read(P/'source-candidate.json')
    accepted = read(P/'analysis.json')['decision']['accepted']
    final_source = candidate if accepted else baseline
    assert custody.census() == final_source == read(P/'source-final.json')
    target = 'crates/litchi-ooxml-common/src/mce/codec.rs'
    assert {n for n in set(baseline)|set(candidate) if baseline.get(n) != candidate.get(n)} == {target}
    original_source = read(P.parent/'change-0712/source-final.json')
    assert original_source[target] == sha(P/'original-codec.rs')
    original_source[target] = baseline[target]
    assert original_source == baseline
    assert sha(P/'baseline-codec.rs') == baseline[target]
    assert sha(P/'candidate-codec.rs') == candidate[target]
    for name, digest in read(P/'constraints.json').items():
        assert sha(ROOT/name) == digest, name
    for filename, folder in [('capture-freeze.json', P), ('oracle-freeze.json', P/'oracle'), ('mechanism-freeze.json', P)]:
        for name, digest in read(P/filename).items():
            assert sha(folder/name) == digest, name
    cleanup = read(P/'cleanup.json')
    expected_paths = ['/home/zhuhe/code/litchi-target-0713', '/home/zhuhe/code/litchi-0713-bin', '/home/zhuhe/code/litchi-0713-fs']
    assert cleanup['owned_paths'] == expected_paths
    assert cleanup['owned_paths_absent'] and all(not Path(n).exists() for n in expected_paths)
    assert len(cleanup['binaries']) == 6
    assert len({r['path'] for r in cleanup['binaries']}) == 6
    for stage in ['baseline', 'candidate']:
        for build in read(P/f'build-{stage}.json'):
            assert build['exit_code'] == 0
            assert build['source_manifest_sha256'] == sha(P/f'source-{stage}.json')
            assert build['log_sha256'] == sha(P/f"build-{stage}-{build['lane']}.log")
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
        assert report['case_count'] == len(report['cases']) == 11
    assert (P/'oracle/baseline/report.json').read_bytes() == (P/'oracle/candidate/report.json').read_bytes()
    assert (P/'oracle/baseline/report.json').read_bytes() == (P.parent/'change-0712/oracle/current/report.json').read_bytes()
    analysis = read(P/'analysis.json')
    assert analysis['verification']['child_receipts_verified'] == 32
    decision = analysis['decision']
    assert decision['deterministic_output_parity_pass']
    failed = [r for r in decision['hard_gates'] if not r['pass']]
    assert len(decision['hard_gates']) == 48
    assert accepted == (not failed)
    assert read(P/'disposition.json')['failed_gates'] == failed
    assert read(P/'pilot-gate.json')['analysis_sha256'] == sha(P/'analysis.json')
    assert read(P/'pilot-gate.json')['status'] == read(P/'disposition.json')['decision'] == ('accepted' if accepted else 'rejected')
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
    with tempfile.TemporaryDirectory(prefix='litchi-0713-audit-') as directory:
        output = Path(directory)/'analysis.json'
        subprocess.run(['python3', '-B', str(P/'analyze.py'), '--output', str(output)], cwd=ROOT, check=True, stdout=subprocess.DEVNULL)
        assert output.read_bytes() == (P/'analysis.json').read_bytes()
    assert read(P/'quality-source.json') == final_source
    quality = read(P/'quality.json')
    assert [r['name'] for r in quality] == ['fmt', 'tests', 'clippy-unaffected-targets', 'clippy-all-libs', 'doctests', 'rustdoc']
    expected_packages = ['litchi-ooxml-common', 'litchi-drawingml', 'litchi-spreadsheet-drawing', 'litchi-docx', 'litchi-xlsx', 'litchi-pptx', 'litchi-xlsb']
    for row in quality:
        command = row['command']
        packages = [command[i+1] for i, value in enumerate(command) if value == '-p']
        expected = [] if row['name'] == 'fmt' else expected_packages[:-2] if row['name'] == 'clippy-unaffected-targets' else expected_packages
        assert packages == expected
        if packages:
            assert '--all-features' in command and '--locked' in command
        if row['name'] in ['tests', 'clippy-unaffected-targets']:
            assert '--all-targets' in command
        if row['name'].startswith('clippy'):
            assert command[-3:] == ['--', '-D', 'warnings']
        if row['name'] == 'clippy-all-libs':
            assert '--lib' in command
        if row['name'] == 'doctests':
            assert '--doc' in command
        if row['name'] == 'rustdoc':
            assert command[:2] == ['env', 'RUSTDOCFLAGS=-D warnings']
    initial = read(P/'quality-initial.json')
    assert [r['exit_code'] for r in initial] == [0, 0, 101]
    assert initial[:2] == quality[:2]
    assert initial[-1]['log_sha256'] == sha(P/'quality-clippy.log')
    second = read(P/'quality-second-attempt.json')
    assert [r['exit_code'] for r in second] == [0, 0, 101]
    assert second[:2] == quality[:2]
    assert second[-1]['log_sha256'] == sha(P/'quality-clippy-direct.log')
    pptx_lint = read(P/'baseline-pptx-clippy.json')
    assert pptx_lint['exit_code'] == 101
    assert pptx_lint['source_manifest_sha256'] == sha(P/'source-baseline.json')
    assert pptx_lint['log_sha256'] == sha(P/'baseline-pptx-clippy.log')
    for name in ['quality-clippy-direct.log', 'baseline-pptx-clippy.log']:
        log = (P/name).read_text()
        assert log.count('error: called `.err().expect()` on a `Result` value') == 3
        assert all(location in log for location in ['opened/tests.rs:464', 'opened/tests.rs:538', 'opened/tests.rs:557'])
    baseline_lint = read(P/'baseline-xlsb-clippy.json')
    assert baseline_lint['exit_code'] == 101
    assert baseline_lint['source_manifest_sha256'] == sha(P/'source-baseline.json')
    assert baseline_lint['log_sha256'] == sha(P/'baseline-xlsb-clippy.log')
    for name in ['quality-clippy.log', 'baseline-xlsb-clippy.log']:
        log = (P/name).read_text()
        assert log.count('error: used `expect()` on a `Result` value') == 4
        for location in ['comments/threaded/tests/mod.rs:271', 'comments/threaded/tests/mod.rs:276', 'comments/threaded/tests/mod.rs:384', 'shared_workbook/tests.rs:180']:
            assert location in log
    for row in quality:
        assert row['exit_code'] == 0
        assert row['source_manifest_sha256'] == sha(P/'quality-source.json')
        assert row['log_sha256'] == sha(P/row['log'])
    for stage in ['baseline', 'candidate']:
        checks = read(P/f'{stage}-checks.json')
        # Baseline has two independently logged commands; candidate has a focused suite.
        rows = checks if isinstance(checks, list) else [checks]
        assert all(row['exit_code'] == 0 for row in rows)
        for row in rows:
            codec_digest = row.get('codec_sha256', row.get('source_sha256'))
            assert codec_digest == (baseline if stage == 'baseline' else candidate)[target]
            log = row['name']+'.log' if stage == 'baseline' else 'candidate-search-tests.log'
            assert row['log_sha256'] == sha(P/log)
            command = row['command']
            assert command[:4] in [['cargo', 'test', '-p', 'litchi-ooxml-common'], ['cargo', 'clippy', '-p', 'litchi-ooxml-common']]
            assert '--locked' in command and '--target-dir' in command
            if command[1] == 'test':
                expected_filter = 'mce::codec::search_equivalence_tests' if stage == 'baseline' else 'mce::codec'
                assert command[-2:] == ['--lib', expected_filter]
                count = 2 if stage == 'baseline' else 3
                assert f'test result: ok. {count} passed; 0 failed;' in (P/log).read_text()
            else:
                assert '--all-targets' in command and command[-3:] == ['--', '-D', 'warnings']
    preparation = read(P/'source-preparation.json')
    for patch, digest in preparation['patch_sha256'].items():
        assert sha(P/patch) == digest
    before = (P/'baseline-codec.rs').read_text()
    old = '''    if needle.is_empty() {
        return Some(0);
    }
    haystack
        .windows(needle.len())
        .position(|window| window == needle)'''
    new = '    memchr::memmem::find(haystack, needle)'
    assert before.count(old) == 1
    assert (before.replace(old, new, 1)+(P/'marker-exhaustion.fragment').read_text()) == (P/'candidate-codec.rs').read_text()
    for left, right, patch in [('original', 'baseline', 'baseline-tests.patch'), ('baseline', 'candidate', 'candidate.patch')]:
        expected = ''.join(difflib.unified_diff((P/(left+'-codec.rs')).read_text().splitlines(True), (P/(right+'-codec.rs')).read_text().splitlines(True), fromfile='a/'+target, tofile='b/'+target))
        assert expected == (P/patch).read_text()
    for label in ['original', 'baseline', 'candidate']:

        assert preparation[label+'_sha256'] == sha(P/(label+'-codec.rs'))
    assert (P/'baseline-codec.rs').read_bytes() == (P/'original-codec.rs').read_bytes() + b'\n' + (P/'search-tests.fragment').read_bytes()
    for row in read(P/'preflight-evidence/results.json'):
        assert row['exit_code'] == 0
        assert row['source_manifest_sha256'] == sha(P/'source-preflight.json')
        assert row['log_sha256'] == sha(P/'preflight-evidence'/(row['name']+'.log'))
    if accepted:
        spec = importlib.util.spec_from_file_location('mechanism0713audit', P/'analyze_mechanism.py')
        mechanism = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mechanism)
        replay = mechanism.analyze(mechanism.load_plan())
        assert (json.dumps(replay, indent=2, sort_keys=True)+'\n').encode() == (P/'mechanism-analysis.json').read_bytes()
        assert replay['callgrind']['children'] == 8 and replay['rss']['children'] == 16
        assert replay['output_parity']['all_outputs_match']
        subprocess.run(['python3', '-B', str(P/'search-attribution.py'), '--check'], cwd=ROOT, check=True, stdout=subprocess.DEVNULL)
        assert read(P/'mechanism-recovery.json')['native_plan_capture_unchanged']
        failed_attempt = P/'mechanism-incomplete-attempt'
        assert not list(failed_attempt.glob('*.receipt.json'))
        assert len(list(failed_attempt.glob('mechanism-profile-baseline-A1-generated-edit.*'))) == 13
        assert read(P/'disposition.json')['production'] == 'candidate retained'
    if not accepted:
        assert not (P/'mechanism-analysis.json').exists()
        assert not list(P.glob('*.callgrind*'))
    print('PASS: 32 children, 11-case parity, decision, final source, quality, six gates, docs, cleanup and exact replay')

if __name__ == '__main__':
    main()
