"""Offline replay of preservation repair, quality receipts, and artifact admission."""
import argparse
import hashlib
import json
import re
from pathlib import Path
import subprocess
import sys

P = Path(__file__).resolve().parent
ROOT = P.parents[3]


def read(path):
    return json.loads(path.read_text())


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def descriptor(value):
    path = P / value['path']
    assert path.is_file() and not path.is_symlink(), path
    assert sha(path) == value['sha256'], path
    if 'bytes' in value:
        assert path.stat().st_size == value['bytes'], path
    return path


def command(label):
    folder = P / 'commands' / label
    result = read(folder / 'result.json')
    started = read(folder / 'started.json')
    assert all(result[k] == v for k, v in started.items())
    assert result['label'] == label and result['started'] <= result['ended']
    assert result['driver_sha256'] == sha(P / 'run.py')
    assert result['plan_sha256'] == sha(P / 'plan.json')
    assert result['log_sha256'] == sha(folder / 'output.log')
    return result, read(descriptor(result['source']))


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--final', action='store_true')
    args = parser.parse_args()
    plan, outcome = read(P / 'plan.json'), read(P / 'outcome.json')
    assert outcome['schema'] == 'litchi.performance.0818.outcome.v1'
    assert outcome['status'] == 'preservation_repaired_and_artifacts_admitted'
    assert outcome['timing_reports'] == 0 and outcome['timing_samples'] == 0
    assert set(outcome['source_allowlist']) == {
        'crates/litchi-docx/src/package/codec.rs',
        'crates/litchi-docx/tests/real_file_preservation.rs',
    }
    for group in ('architecture', 'unrelated'):
        for name, digest in plan[group].items():
            assert sha(ROOT / name) == digest, name
    assert sha(ROOT / 'Cargo.lock') == plan['root_lock_sha256']
    assert sha(ROOT / 'tools/perf-baseline/Cargo.lock') == plan['tool_lock_sha256']
    for name, identity in plan['admission']['real_inputs'].items():
        assert sha(ROOT / name) == identity['sha256']
        assert (ROOT / name).stat().st_size == identity['bytes']
    # The prior failed admission is immutable, including every sealed payload.
    old = ROOT / 'docs/performance/results/change-0817/seal.json'
    for name, digest in read(old)['files'].items():
        # This batch changes indexed summaries, but no prior packet or test.
        if '/results/change-0817/' in name or name.startswith('tools/perf-baseline/'):
            assert sha(ROOT / name) == digest, name
    preflight = read(P / 'preflight/receipt.json')
    assert preflight['exit_code'] == 1
    assert preflight['auditor']['sha256'] == sha(P / 'artifact_audit.py')
    for key in ('report', 'log', 'artifacts_manifest', 'frozen_0817_auditor'):
        item = preflight[key]
        assert sha(Path(item['path'])) == item['sha256']
    preflight_audit = read(P / 'preflight/artifact-audit.json')
    assert preflight_audit['ok'] is False and len(preflight_audit['errors']) == 3
    assert all(error.startswith('real-000-docx:') for error in preflight_audit['errors'])
    assert sum(case['ok'] for case in preflight_audit['cases']) == 5
    baseline, old_source = command('baseline-regression')
    assert baseline['exit_code'] == 101
    log = (P / 'commands/baseline-regression/output.log').read_text()
    assert 'test docx_ordinary_save_preserves_document_relationship_bytes ... FAILED' in log
    assert 'an untouched document relationship member must retain exact bytes and order' in log
    assert '0 passed; 1 failed' in log
    test = 'crates/litchi-docx/tests/real_file_preservation.rs'
    codec = 'crates/litchi-docx/src/package/codec.rs'
    assert sha(P / 'baseline-test.rs') == old_source[test]
    original = subprocess.check_output(['git', 'show', plan['base'] + ':' + codec], cwd=ROOT)
    assert hashlib.sha256(original).hexdigest() == old_source[codec]
    expected_commands = outcome['successful_commands']
    assert expected_commands == ['fmt', 'check', 'tests', 'clippy', 'rustdoc', 'boundaries', 'build-exporter', 'export-artifacts']
    final_source = None
    for label in expected_commands:
        receipt, source = command(label)
        assert receipt['exit_code'] == 0, label
        if final_source is None:
            final_source = source
        assert source == final_source, label
    assert set(final_source) == set(old_source)
    changed = {name for name in final_source if final_source[name] != old_source[name]}
    assert codec in changed and changed <= set(outcome['source_allowlist'])
    for name, digest in final_source.items():
        assert sha(ROOT / name) == digest, name
    changed_tracked = set(subprocess.check_output(['git', 'diff', '--name-only', plan['base'], '--', 'crates', 'Cargo.toml', 'Cargo.lock', 'rustfmt.toml', 'tools/perf-baseline'], cwd=ROOT, text=True).splitlines())
    assert changed_tracked <= set(outcome['source_allowlist'])
    test_log = (P / 'commands/tests/output.log').read_text()
    suites = re.findall(r'test result: (\w+)\. (\d+) passed; (\d+) failed; (\d+) ignored;', test_log)
    assert len(suites) == 66 and all(row[0] == 'ok' and row[2] == '0' for row in suites)
    assert sum(int(row[1]) for row in suites) == 1938
    assert sum(int(row[3]) for row in suites) == 32
    assert 'test docx_ordinary_save_preserves_document_relationship_bytes ... ok' in test_log
    assert 'test docx_changed_document_relationships_are_serialized ... ok' in test_log
    assert outcome['tests'] == read(P / 'test-summary.json')
    admission = read(descriptor(outcome['admission_receipt']))
    assert admission['exit_code'] == 0 and admission['started'] <= admission['ended']
    assert admission['auditor']['sha256'] == outcome['auditor_sha256']
    assert admission['manifest'] == outcome['manifest'] and admission['report'] == outcome['audit']
    descriptor(admission['log'])
    manifest = read(descriptor(outcome['manifest']))
    assert len(manifest['cases']) == 6
    policy_count = 0
    real_sources = {}
    for case in manifest['cases']:
        outputs = [row['output'] for row in case['policy_outputs']] + [case['stream_output']['output']]
        assert len(outputs) == 5 and len({x['sha256'] for x in outputs}) == 1
        assert case['edit_admitted'] and case['edit_outcome'] == 'admitted'
        for value in [case['source_archive']] + outputs:
            path = P / 'artifacts' / value['path']
            assert path.is_file() and sha(path) == value['sha256'] and path.stat().st_size == value['bytes']
        if case['origin'] == 'caller-named-real-file':
            real_sources[case['format']] = case['source_archive']['sha256']
        policy_count += len(outputs)
    assert policy_count == 30 and set(real_sources.values()) == {x['sha256'] for x in plan['admission']['real_inputs'].values()}
    audit_path = descriptor(outcome['audit'])
    audit = read(audit_path)
    assert audit['ok'] is True and audit['errors'] == []
    assert len(audit['cases']) == 6 and all(c['ok'] for c in audit['cases'])
    assert outcome['auditor_sha256'] == sha(P / 'artifact_audit.py')
    subprocess.run([sys.executable, '-B', str(P / 'artifact_audit.py'), '--check', str(audit_path)], check=True, cwd=ROOT, stdout=subprocess.DEVNULL)
    descriptor(outcome['zip_preservation'])
    subprocess.run([sys.executable, '-B', str(P / 'preservation.py'), '--check'], check=True, cwd=ROOT, stdout=subprocess.DEVNULL)
    binary = outcome['binary']
    if Path(binary['path']).exists():
        assert sha(Path(binary['path'])) == binary['sha256']
    else:
        cleanup = read(P / 'cleanup.json')
        assert cleanup['binary'] == binary and cleanup['binary_verified_before_removal'] is True
    if args.final:
        cleanup = read(P / 'cleanup.json')
        assert cleanup['removed'] and all(not Path(row['path']).exists() for row in cleanup['removed'])
        assert {row['path'] for row in cleanup['removed']} == {plan['target'], plan['scratch']}
        assert not list(P.rglob('__pycache__'))
    print('0818 replay PASS: baseline regression failed, repaired quality passed, 6 corpora / 30 outputs admitted; no timing claim')


if __name__ == '__main__':
    main()
