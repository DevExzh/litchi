"""Reconcile source replay, serial receipts, reports and the optional evidence seal."""
import argparse
import datetime
import hashlib
import json
import os
from pathlib import Path
import re
import subprocess
import tempfile

from run import HERE, REPO, sha
from checks import COMMANDS


def read(path):
    return json.loads(path.read_text())


def verify(sealed=False):
    plan = read(HERE / 'plan.json')
    manifests = {stage: read(HERE / stage / 'source-manifest.json')
                 for stage in ('baseline', 'candidate')}
    assert manifests['baseline'].keys() == manifests['candidate'].keys()
    differences = [p for p in manifests['baseline']
                   if manifests['baseline'][p] != manifests['candidate'][p]]
    assert differences == ['crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/scan.rs']
    for name, digest in manifests['baseline'].items():
        assert sha(REPO / name) == digest, name
    disposition = read(HERE / 'disposition.json')
    assert disposition['status'] == 'rejected_and_reverted'
    assert disposition['final_source_manifest_sha256'] == sha(HERE / 'baseline/source-manifest.json')
    assert disposition['comparison_sha256'] == sha(HERE / 'comparison.json')
    assert disposition['plan_sha256'] == sha(HERE / 'plan.json')
    assert disposition['production_change_retained'] is False
    for name, digest in read(HERE / 'adr-manifest.json')['files'].items():
        assert sha(REPO / name) == digest, name
    review = (HERE / 'final-source-review.md').read_text()
    for name in ('crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/scan.rs',
                 'crates/litchi-xlsx/src/raw/worksheet/edit/codec/tests.rs',
                 'tools/perf-baseline/src/lib.rs'):
        assert manifests['candidate'][name] in review
    # Replay tracked patches through a private Git index without altering the
    # checkout. Tests and the noncompact harness are identical in both stages.
    with tempfile.TemporaryDirectory(prefix='litchi-0522-replay-') as directory:
        for stage in manifests:
            env = dict(os.environ, GIT_INDEX_FILE=str(Path(directory) / stage))
            subprocess.run(['git', 'read-tree', plan['revision']], cwd=REPO, env=env, check=True)
            subprocess.run(['git', 'apply', '--cached', str(HERE / stage / 'source.patch')],
                           cwd=REPO, env=env, check=True)
            entries = subprocess.check_output(['git', 'ls-files', '-s', '-z'], cwd=REPO, env=env)
            objects = {}
            for entry in entries.split(b'\0'):
                if entry:
                    meta, name = entry.split(b'\t', 1)
                    objects[name.decode()] = meta.split()[1].decode()
            unique = {objects[name] for name in manifests[stage] if name in objects}
            batch = subprocess.check_output(['git', 'cat-file', '--batch'], cwd=REPO,
                                            input=('\n'.join(sorted(unique)) + '\n').encode())
            hashes = {}
            position = 0
            while position < len(batch):
                end = batch.index(b'\n', position)
                oid, kind, size = batch[position:end].split()
                assert kind == b'blob'
                position = end + 1
                data = batch[position:position + int(size)]
                hashes[oid.decode()] = hashlib.sha256(data).hexdigest()
                position += int(size) + 1
            for name, digest in manifests[stage].items():
                assert name in objects and hashes[objects[name]] == digest, (stage, name)
    records = []
    stage_counts = {}
    for stage in manifests:
        folder = HERE / stage
        receipts = sorted(folder.glob('*.receipt.json'))
        stage_counts[stage] = len(receipts)
        assert len(receipts) == (29 if stage == 'baseline' else 40)
        for path in receipts:
            receipt = read(path)
            assert receipt['exit_code'] == 0, path
            assert receipt['source_manifest_sha256'] == sha(folder / 'source-manifest.json')
            working_stage = 'candidate' if stage == 'baseline' and path.name.startswith('native-r2-') else stage
            assert receipt['working_source_manifest_sha256'] == sha(HERE / working_stage / 'source-manifest.json')
            assert receipt['script_sha256'] == sha(HERE / 'run.py')
            assert receipt['plan_sha256'] == sha(HERE / 'plan.json')
            for name, digest in receipt['artifacts'].items():
                assert Path(name).name == name and sha(folder / name) == digest, name
            records.append(receipt)
    for name, command in COMMANDS:
        receipt = read(HERE / 'candidate' / ('check-' + name + '.receipt.json'))
        expected = ['env', 'CARGO_TARGET_DIR=' + plan['owned_paths'][1],
                    'CARGO_BUILD_JOBS=2', 'CARGO_INCREMENTAL=0', 'RUSTDOCFLAGS=-D warnings'] + command
        assert receipt['command'] == expected, name
    final_check = read(HERE / 'baseline/check-final-codec.receipt.json')
    assert final_check['command'] == ['env', 'CARGO_BUILD_JOBS=2', 'CARGO_INCREMENTAL=0',
        'cargo', 'test', '--locked', '-p', 'litchi-xlsx', '--lib',
        'raw::worksheet::edit::codec::tests', '--target-dir', plan['owned_paths'][1]]
    final_log = (HERE / 'baseline/check-final-codec.stdout').read_text()
    assert sum(map(int, re.findall(r'test result: ok\. (\d+) passed;', final_log))) == 19
    quality = read(HERE / 'quality-summary.json')
    assert quality['status'] == 'pass' and len(quality['checks']) == len(COMMANDS) == 12
    for row in quality['checks']:
        path = HERE / 'candidate' / row['name']
        assert sha(path) == row['receipt_sha256'] and row['exit_code'] == 0
        log = path.with_name(path.name.replace('.receipt.json', '.stdout')).read_text()
        assert row['executed_tests'] == sum(map(int, re.findall(r'test result: ok\. (\d+) passed;', log)))
    assert quality['executed_tests'] == sum(row['executed_tests'] for row in quality['checks']) and quality['executed_tests'] > 1278
    focused = read(HERE / 'baseline-focused.json')
    assert focused['exit_code'] == 0
    for name, digest in focused['source_hashes'].items():
        assert manifests['baseline'][name] == digest
    harness = read(HERE / 'baseline-harness.json')
    assert harness['exit_code'] == 0
    for name, digest in harness['source_hashes'].items():
        assert manifests['baseline'][name] == digest
    records.sort(key=lambda r: r['start_utc'])
    for receipt in records:
        assert datetime.datetime.fromisoformat(receipt['start_utc']) < datetime.datetime.fromisoformat(receipt['end_utc'])
        assert receipt['seconds'] > 0
    assert all(a['end_utc'] <= b['start_utc'] for a, b in zip(records, records[1:]))
    native_order = []
    for stage in manifests:
        for path in (HERE / stage).glob('native-*.receipt.json'):
            native_order.append((read(path)['start_utc'], stage, path.name.split('-')[1]))
    native_order.sort()
    blocks = []
    for _, stage, repeat in native_order:
        if not blocks or blocks[-1] != (stage, repeat):
            blocks.append((stage, repeat))
    assert blocks == [('baseline', 'r1'), ('candidate', 'r1'), ('candidate', 'r2'), ('baseline', 'r2')]
    annotations = {p: p.read_bytes() for p in HERE.glob('*/profile-*.txt')}
    with tempfile.TemporaryDirectory(prefix='litchi-0522-analysis-') as directory:
        for script, filename in [('analyze.py', 'comparison.json'),
                                 ('analyze_profiles.py', 'profile-analysis.json')]:
            output = Path(directory) / filename
            subprocess.run(['python3', '-B', str(HERE / script), str(output)], cwd=REPO, check=True)
            assert output.read_bytes() == (HERE / filename).read_bytes(), filename
    assert len(annotations) == 16
    assert all(path.read_bytes() == content for path, content in annotations.items())
    negative = read(HERE / 'verifier-tests.json')
    assert negative['status'] == 'pass' and len(negative['checks']) == 4
    assert all(row['rejected'] for row in negative['checks'])
    flags = read(HERE / 'flag-review.json')
    comparison = read(HERE / 'comparison.json')['comparison']
    assert flags['comparison_sha256'] == sha(HERE / 'comparison.json')
    assert flags['adverse_flag_count'] == len(comparison['adverse_flags_over_five_percent'])
    assert flags['same_build_drift_count'] == len(comparison['same_build_drift_over_five_percent'])
    assert len(flags['flags']) == flags['adverse_flag_count']
    for retained, reviewed in zip(comparison['adverse_flags_over_five_percent'], flags['flags']):
        assert all(reviewed[k] == value for k, value in retained.items())
        assert reviewed['review']
    cleanup = read(HERE / 'cleanup.json')
    assert not cleanup['accessible_process_references']
    assert all(not Path(name).exists() for name in plan['owned_paths'])
    assert cleanup['python_cache_absent']
    assert not list(HERE.rglob('__pycache__'))
    seal_count = None
    if sealed:
        expected = {}
        for line in (HERE / 'SHA256SUMS').read_text().splitlines():
            digest, name = line.split('  ', 1)
            assert name not in expected and not Path(name).is_absolute() and '..' not in Path(name).parts
            assert name != 'SHA256SUMS'
            expected[name] = digest
        actual = {str(p.relative_to(HERE)): sha(p) for p in HERE.rglob('*')
                  if p.is_file() and p.name != 'SHA256SUMS'}
        assert expected == actual
        assert not any(p.is_symlink() for p in HERE.rglob('*'))
        seal_count = len(expected)
    return dict(status='pass', native_samples=1640, allocator_samples=40, profile_samples=8,
                stage_receipts=stage_counts, serialized_intervals=len(records),
                source_replay=True, measured_source_differences=differences,
                disposition='rejected_and_reverted', final_source='baseline',
                final_codec_tests=19, exact_report_replay=True,
                exact_annotation_replay=True, negative_vectors=4, owned_paths_absent=True,
                seal_entries=seal_count)


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--sealed', action='store_true')
    parser.add_argument('--output', type=Path)
    args = parser.parse_args()
    report = verify(args.sealed)
    if args.output:
        args.output.write_text(json.dumps(report, indent=2) + '\n')
    print(json.dumps(report))
