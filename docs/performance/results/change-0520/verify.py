"""Verify retained 0520 custody, serialized execution, analyses and optional seal."""
import argparse
import datetime
import json
from pathlib import Path
import subprocess
import tempfile

from capture import HERE, REPO, sha, source_check


def verify(sealed):
    source_check()
    for filename, digest in json.loads((HERE / 'adr-manifest.json').read_text())['files'].items():
        assert sha(REPO / filename) == digest, filename
    build = json.loads((HERE / 'build.json').read_text())
    plan = json.loads((HERE / 'plan.json').read_text())
    assert build['exit_code'] == 0 and build['command'] == plan['build']
    assert build['source_manifest_sha256'] == sha(HERE / 'source-manifest.json')
    receipts = sorted(HERE.glob('*.receipt.json'))
    assert len(receipts) == 8
    records = [build]
    for path in receipts:
        r = json.loads(path.read_text())
        assert r['exit_code'] == 0
        assert r['binary_sha256'] == build['binary_sha256']
        assert r['source_manifest_sha256'] == sha(HERE / 'source-manifest.json')
        assert r['plan_sha256'] == sha(HERE / 'plan.json')
        assert r['script_sha256'] == sha(HERE / 'capture.py')
        for f, h in r['artifacts'].items():
            assert Path(f).name == f and sha(HERE / f) == h, f
        records.append(r)
    records.sort(key=lambda r: r['start_utc'])
    assert records[0] == build
    for record in records:
        start = datetime.datetime.fromisoformat(record['start_utc'])
        end = datetime.datetime.fromisoformat(record['end_utc'])
        assert start <= end and record['seconds'] > 0
    assert all(a['end_utc'] <= b['start_utc'] for a, b in zip(records, records[1:]))
    annotations = {p.name: p.read_bytes() for p in HERE.glob('profile-*.txt')}
    replays = []
    with tempfile.TemporaryDirectory(prefix='litchi-0520-verify-') as directory:
        for script, filename in [('analyze.py', 'analysis.json'), ('analyze_profiles.py', 'profile-analysis.json')]:
            output = Path(directory) / filename
            result = subprocess.run(['python3', '-B', str(HERE/script), str(output)],
                                    cwd=REPO, capture_output=True, text=True)
            assert result.returncode == 0, (script, result.stdout, result.stderr)
            assert output.read_bytes() == (HERE / filename).read_bytes(), filename
            replays.append(dict(script=script, sha256=sha(HERE/script), report=filename,
                                report_sha256=sha(output), exact_replay=True))
    assert all((HERE / filename).read_bytes() == content for filename, content in annotations.items())
    cleanup = json.loads((HERE / 'cleanup.json').read_text())
    assert cleanup['owned_target_absent'] and cleanup['python_cache_absent']
    assert not cleanup['accessible_process_references']
    assert all(not Path(p).exists() for p in cleanup['removed'])
    checks = json.loads((HERE / 'checks.json').read_text())
    assert len(checks['checks']) == 2 and all(r['exit_code'] == 0 for r in checks['checks'])
    negative = json.loads((HERE / 'verifier-tests.json').read_text())
    assert negative['status'] == 'pass' and negative['valid_control_and_restoration_pass']
    assert len(negative['checks']) == 4 and all(r['rejected'] for r in negative['checks'])
    assert negative['temporary_directory_removed']
    seal_entries = None
    if sealed:
        expected = {}
        for line in (HERE / 'SHA256SUMS').read_text().splitlines():
            digest, filename = line.split('  ', 1)
            assert filename not in expected
            assert Path(filename).name == filename and filename != 'SHA256SUMS'
            expected[filename] = digest
        actual = {p.name: sha(p) for p in HERE.iterdir() if p.is_file() and p.name != 'SHA256SUMS'}
        assert expected == actual
        assert all(p.is_file() and not p.is_symlink() for p in HERE.iterdir())
        seal_entries = len(expected)
    return dict(status='pass', native_samples=400, profile_samples=4,
                serialized_build_and_capture_intervals=len(records),
                exact_report_replays=replays, annotations_exact=True,
                owned_target_absent=True, negative_vectors_rejected=4,
                seal_entries=seal_entries,
                limitations='No new full-workspace tests, physical I/O, allocator/RSS, native Office, cold/range or scaling claim.')


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--sealed', action='store_true')
    parser.add_argument('--output', type=Path)
    args = parser.parse_args()
    report = verify(args.sealed)
    if args.output:
        args.output.write_text(json.dumps(report, indent=2) + '\n')
    print(json.dumps(report))
