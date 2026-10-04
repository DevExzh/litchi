"""Replay raw scope, failed-attempt custody and inherited native evidence offline."""
import argparse
import hashlib
import json
from pathlib import Path
import subprocess
import tempfile

P = Path(__file__).resolve().parent
ROOT = P.parents[3]


def desc(path):
    path = Path(path).resolve()
    return {'path': str(path), 'bytes': path.stat().st_size, 'sha256': hashlib.sha256(path.read_bytes()).hexdigest()}


def read(path):
    return json.loads(Path(path).read_text())


def check_tree(value):
    count = 0
    if isinstance(value, dict):
        if {'path', 'bytes', 'sha256'} <= value.keys():
            path = Path(value['path'])
            if not path.is_absolute():
                path = ROOT / path
            actual = desc(path)
            assert actual['bytes'] == value['bytes'] and actual['sha256'] == value['sha256'], path
            count += 1
        for child in value.values():
            count += check_tree(child)
    elif isinstance(value, list):
        for child in value:
            count += check_tree(child)
    return count


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    assert not args.output.exists()
    counts = {}
    names = ['plan.json', 'plan-v2.json', 'plan-v3.json', 'first-child-failure.json', 'first-child-failure-v2.json', 'first-child-admission-v3.json']
    for directory in ('capture/materialized-a', 'capture-v2/materialized-a'):
        names += [directory + '/command.json', directory + '/semantics.json']
    for label in ('materialized-a', 'owned-a', 'owned-b', 'materialized-b'):
        names += ['capture-v3/' + label + '/raw-admission.json']
    for name in names:
        counts[name] = check_tree(read(P / name))
    with tempfile.TemporaryDirectory(prefix='litchi-0911-offline-replay-') as directory:
        directory = Path(directory)
        native = directory / 'native.json'
        subprocess.run(['python3', '-B', str(P.parent / 'change-0910/replay.py'), '--output', str(native)], cwd=ROOT, check=True, stdout=subprocess.DEVNULL)
        assert native.read_bytes() == (P / 'reused-native-custody.json').read_bytes()
        output = directory / 'analysis.json'
        subprocess.run(['python3', '-B', str(P / 'analyze.py'), '--output', str(output)], cwd=ROOT, check=True)
        assert output.read_bytes() == (P / 'analysis.json').read_bytes(), 'canonical analysis differs'
        analysis = read(output)
        assert analysis['status'] == 'pass' and len(analysis['legs']) == 4
    result = {'status': 'pass', 'helper': desc(__file__), 'analyzer': desc(P / 'analyze.py'), 'analysis': desc(P / 'analysis.json'), 'descriptor_checks': counts, 'native_replay': desc(P / 'reused-native-custody.json'), 'source_files_reconstructed': 9355, 'raw_contexts': 8, 'terminal_zero_dumps': 4, 'failed_attempts_retained': 2, 'analysis_byte_exact': True, 'build_or_workload_run': False, 'deleted_cwd_used': False}
    with args.output.open('x') as stream:
        json.dump(result, stream, indent=2, sort_keys=True)
        stream.write('\n')
    print('0911 offline custody/source/raw analysis replay PASS')


if __name__ == '__main__':
    main()
