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
    assert differences == ['crates/litchi-xlsx/src/cell_values/validation.rs']
    for name, digest in manifests['candidate'].items():
        assert sha(REPO / name) == digest, name
    for name, digest in read(HERE / 'adr-manifest.json')['files'].items():
        assert sha(REPO / name) == digest, name
    review = (HERE / 'final-source-review.md').read_text()
    for name in ('crates/litchi-xlsx/src/cell_values/validation.rs',
                 'crates/litchi-xlsx/src/cell_values/validation_borrow_tests.rs',
                 'tools/perf-baseline/src/lib.rs',
                 'tools/perf-baseline/src/xlsx_commit_metrics_tests.rs'):
        assert sha(REPO / name) in review
    # Replay tracked patches through a private Git index, without altering the
    # checkout. The new differential test is identical in both manifests.
    with tempfile.TemporaryDirectory(prefix='litchi-0521-replay-') as directory:
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
                if name in objects:
                    assert hashes[objects[name]] == digest, (stage, name)
                else:
                    assert name == 'crates/litchi-xlsx/src/cell_values/validation_borrow_tests.rs'
                    assert sha(REPO / name) == digest
    records = []
    stage_counts = {}
    for stage in manifests:
        folder = HERE / stage
        receipts = sorted(folder.glob('*.receipt.json'))
        stage_counts[stage] = len(receipts)
        assert len(receipts) == (24 if stage == 'baseline' else 35)
        for path in receipts:
            receipt = read(path)
            assert receipt['exit_code'] == 0, path
            assert receipt['source_manifest_sha256'] == sha(folder / 'source-manifest.json')
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
    quality = read(HERE / 'quality-summary.json')
    assert quality['status'] == 'pass' and len(quality['checks']) == 11
    for row in quality['checks']:
        path = HERE / 'candidate' / row['name']
        assert sha(path) == row['receipt_sha256'] and row['exit_code'] == 0
        log = path.with_name(path.name.replace('.receipt.json', '.stdout')).read_text()
        assert row['executed_tests'] == sum(map(int, re.findall(r'test result: ok\. (\d+) passed;', log)))
    assert quality['executed_tests'] == sum(row['executed_tests'] for row in quality['checks']) == 1278
    corrected = read(HERE / 'baseline-focused-corrected.json')
    assert corrected['exit_code'] == 0
    assert read(HERE / 'baseline-focused.json')['exit_code'] == 101
    for name, digest in corrected['source_hashes'].items():
        assert manifests['baseline'][name] == digest
    records.sort(key=lambda r: r['start_utc'])
    for receipt in records:
        assert datetime.datetime.fromisoformat(receipt['start_utc']) < datetime.datetime.fromisoformat(receipt['end_utc'])
        assert receipt['seconds'] > 0
    assert all(a['end_utc'] <= b['start_utc'] for a, b in zip(records, records[1:]))
    annotations = {p: p.read_bytes() for p in HERE.glob('*/profile-*.txt')}
    with tempfile.TemporaryDirectory(prefix='litchi-0521-analysis-') as directory:
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
    assert flags['adverse_flag_count'] == len(comparison['adverse_flags_over_five_percent']) == 23
    assert flags['same_build_drift_count'] == len(comparison['same_build_drift_over_five_percent']) == 52
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
    return dict(status='pass', native_samples=1400, allocator_samples=40, profile_samples=8,
                stage_receipts=stage_counts, serialized_intervals=len(records),
                source_replay=True, source_differences=differences, exact_report_replay=True,
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
