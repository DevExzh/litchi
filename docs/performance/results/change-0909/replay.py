"""Offline reconstruction and custody replay; never builds or runs the Rust probe."""
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


def check(row):
    path = Path(row['path'])
    if not path.is_absolute():
        path = ROOT / path
    assert path.is_file(), path
    assert path.stat().st_size == row['bytes'] and sha(path) == row['sha256'], path


def load(name):
    return json.loads((P / name).read_text())


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--attempt', required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    assert not args.output.exists()
    source = load(f'source-{args.attempt}.json')
    base = load('source-base.json')
    check(source['base'])
    check(source['overlay_patch'])
    archive = subprocess.check_output(['git', 'archive', base['commit'], '--', *base['archive_paths']], cwd=ROOT)
    assert hashlib.sha256(archive).hexdigest() == base['archive_sha256']
    with tempfile.TemporaryDirectory(prefix='litchi-0909-source-replay-') as directory:
        target = Path(directory)
        with tarfile.open(fileobj=io.BytesIO(archive)) as stream:
            stream.extractall(target, filter='data')
        for name, row in base['lock_overlays'].items():
            data = (ROOT / row['artifact']).read_bytes()
            assert hashlib.sha256(data).hexdigest() == row['sha256']
            (target / name).write_bytes(data)
        (target / '.owned-0909').write_text(base['commit'] + '\n')
        subprocess.run(['git', 'apply', str(Path(source['overlay_patch']['path']).resolve())], cwd=target, check=True)
        actual = [{'path': str(f.relative_to(target)), 'bytes': f.stat().st_size, 'sha256': sha(f)} for f in sorted(target.rglob('*')) if f.is_file()]
        assert actual == source['files'], 'reconstructed source inventory differs'
    build = load(f'build-{args.attempt}/command.json')
    capture = load(f'capture-{args.attempt}/command.json')
    for receipt in [build, capture]:
        assert receipt['exit_code'] == 0
        check(receipt['source'])
        check(receipt['helper'])
        for row in receipt['logs']:
            check(row)
    # The build path is a lexical custody identity after cleanup. Its retained
    # executable copy is validated through the explicit relocation receipt.
    relocation = load('binary.json')
    check(relocation['retained'])
    assert relocation['retained']['sha256'] == build['binary']['sha256'] == capture['binary']['sha256']
    for row in capture['artifacts']:
        check(row)
    report = load(f'capture-{args.attempt}/artifacts/report.json')
    assert report['status'] == 'probe_passed_pending_independent_raw_census'
    census = load('census-runs.json')
    replayed = []
    with tempfile.TemporaryDirectory(prefix='litchi-0909-census-replay-') as directory:
        for index, item in enumerate(census['runs']):
            assert item['exit_code'] == 0
            check(item['report'])
            argv = list(item['argv'])
            destination = Path(directory) / f'{index}.json'
            argv[argv.index('--report') + 1] = str(destination)
            subprocess.run(argv, cwd=ROOT, check=True, stdout=subprocess.DEVNULL)
            assert destination.read_bytes() == Path(item['report']['path']).read_bytes()
            replayed.append(item['label'])
        analysis = Path(directory) / 'analysis.json'
        subprocess.run(['python3', '-B', str(P / 'analyze.py'), '--output', str(analysis)], cwd=ROOT, check=True, stdout=subprocess.DEVNULL)
        assert analysis.read_bytes() == (P / 'analysis.json').read_bytes()
    result = {'status': 'pass', 'source_file_count': len(actual), 'source_manifest': f'source-{args.attempt}.json', 'capture_artifacts_checked': len(capture['artifacts']), 'census_byte_exact_replays': replayed, 'deleted_source_root_used': False, 'build_or_workload_run': False}
    with args.output.open('x') as stream:
        json.dump(result, stream, indent=2)
        stream.write('\n')
    print('0909 offline source/custody/census replay passed')


if __name__ == '__main__':
    main()
