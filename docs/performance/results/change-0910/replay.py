"""Reconstruct frozen source and replay custody/censuses/statistics, without Rust execution."""
import argparse
import hashlib
import io
import json
from pathlib import Path
import subprocess
import tarfile
import tempfile

P = Path(__file__).resolve().parent
ROOT = P.parents[3]


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def resolve(path):
    path = Path(path)
    return path if path.is_absolute() else ROOT / path


def check(row):
    path = resolve(row['path'])
    assert path.is_file() and path.stat().st_size == row['bytes'] and sha(path) == row['sha256'], path
    return path


def load(name):
    return json.loads((P / name).read_text())


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--attempt', default='v1')
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    assert not args.output.exists()
    source = load(f'source-{args.attempt}.json')
    base = load('source-base.json')
    check(source['base'])
    patch = check(source['overlay_patch'])
    archive = subprocess.check_output(['git', 'archive', base['commit'], '--', *base['archive_paths']], cwd=ROOT)
    assert hashlib.sha256(archive).hexdigest() == base['archive_sha256']
    with tempfile.TemporaryDirectory(prefix='litchi-0910-source-replay-') as directory:
        target = Path(directory)
        with tarfile.open(fileobj=io.BytesIO(archive)) as stream:
            stream.extractall(target, filter='data')
        for name, row in base['lock_overlays'].items():
            data = (ROOT / row['artifact']).read_bytes()
            assert hashlib.sha256(data).hexdigest() == row['sha256']
            (target / name).write_bytes(data)
        (target / '.owned-0910').write_text(base['commit'] + '\n')
        subprocess.run(['git', 'apply', str(patch)], cwd=target, check=True)
        actual = [{'path': str(f.relative_to(target)), 'bytes': f.stat().st_size, 'sha256': sha(f)} for f in sorted(target.rglob('*')) if f.is_file()]
        assert actual == source['files'], 'reconstructed source inventory differs'
    build = load(f'build-{args.attempt}/command.json')
    relocation = load('binary.json')
    check(relocation['retained'])
    assert relocation['retained']['sha256'] == build['binary']['sha256']
    receipts = [build] + [load(f'quality-{args.attempt}/{kind}/command.json') for kind in ('fmt', 'clippy')]
    legs = load('capture-plan.json')['legs']
    receipts += [load(f"capture/{leg['label']}/command.json") for leg in legs]
    artifacts = 0
    for receipt in receipts:
        assert receipt['exit_code'] == 0
        check(receipt['source'])
        check(receipt['helper'])
        if 'binary' in receipt:
            assert receipt['binary']['sha256'] == relocation['retained']['sha256']
        if 'plan' in receipt:
            check(receipt['plan'])
        for row in receipt['logs']:
            check(row)
        for row in receipt.get('artifacts', []):
            check(row)
            artifacts += 1
        if 'process_metrics' in receipt:
            check(receipt['process_metrics'])
    replayed = []
    with tempfile.TemporaryDirectory(prefix='litchi-0910-offline-replay-') as directory:
        for index, item in enumerate(load('census-runs.json')['runs']):
            assert item['exit_code'] == 0
            original = check(item['report'])
            check(item['helper'])
            argv = list(item['argv'])
            destination = Path(directory) / f'{index}.json'
            argv[argv.index('--report') + 1] = str(destination)
            subprocess.run(argv, cwd=ROOT, check=True, stdout=subprocess.DEVNULL)
            assert destination.read_bytes() == original.read_bytes()
            census = json.loads(destination.read_text())
            assert census['status'] == 'pass' and len(census['unchanged_members']) == 437
            assert sorted(census['changed_members']['names']) == item['changed_members']
            physical = {side: [r['name'] for r in sorted(census['members'], key=lambda r: r[side]['physical']['local_header_offset'])] for side in ('source', 'output')}
            assert physical['source'] == physical['output']
            replayed.append(item['label'])
        analysis = Path(directory) / 'analysis.json'
        subprocess.run(['python3', '-B', str(P / 'analyze.py'), '--output', str(analysis)], cwd=ROOT, check=True, stdout=subprocess.DEVNULL)
        assert analysis.read_bytes() == (P / 'analysis.json').read_bytes()
    result = {'status': 'pass', 'source_file_count': len(actual), 'source_manifest': f'source-{args.attempt}.json', 'capture_artifacts_checked': artifacts, 'legs': len(legs), 'census_byte_exact_replays': replayed, 'deleted_source_root_used': False, 'build_or_workload_run': False}
    with args.output.open('x') as stream:
        json.dump(result, stream, indent=2)
        stream.write('\n')
    print('0910 offline source/custody/census/statistics replay passed')


if __name__ == '__main__':
    main()
