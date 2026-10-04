"""Replay retained perf rejection and GDB reachability offline, with source custody."""
import argparse
import gzip
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
                candidates = [base / path for base in (ROOT, P) if (base / path).is_file()]
                assert len(candidates) == 1, (path, candidates)
                path = candidates[0]
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
    names = ['plan.json', 'plan-gdb.json', 'symbol-ranges.json', 'first-child-failure.json', 'dispatch-analysis.json', 'dispatch-checks.json']
    perf_folder = P / 'capture/materialized-a'
    names += [str(f.relative_to(P)) for f in sorted(perf_folder.glob('*.json'))]
    for kind in ('materialized', 'owned'):
        names += [f'dispatch/{kind}/command.json', f'dispatch/{kind}/semantics.json']
    counts = {name: check_tree(read(P / name)) for name in names}
    plan = read(P / 'plan.json')
    decodes = []
    for name in ('decode', 'mappings', 'header', 'buildids', 'sample-debug'):
        receipt = read(perf_folder / (name + '.json'))
        result = subprocess.run(receipt['argv'], cwd=ROOT, env=plan['environment'], capture_output=True)
        assert result.returncode == receipt['exit_code'] == 0
        stdout = perf_folder / ('script.txt.gz' if name == 'decode' else name + '.stdout.gz')
        stderr = perf_folder / ('decode.stderr.txt' if name == 'decode' else name + '.stderr.txt')
        assert result.stdout == gzip.decompress(stdout.read_bytes()), name
        assert result.stderr == stderr.read_bytes(), name
        decodes.append(name)
    with tempfile.TemporaryDirectory(prefix='litchi-0912-offline-replay-') as directory:
        directory = Path(directory)
        native = directory / 'native.json'
        subprocess.run(['python3', '-B', str(P.parent / 'change-0910/replay.py'), '--output', str(native)], cwd=ROOT, check=True, stdout=subprocess.DEVNULL)
        assert native.read_bytes() == (P / 'reused-native-custody.json').read_bytes()
        dispatch = directory / 'dispatch.json'
        subprocess.run(['python3', '-B', str(P / 'verify_dispatch.py'), '--output', str(dispatch)], cwd=ROOT, check=True, stdout=subprocess.DEVNULL)
        assert dispatch.read_bytes() == (P / 'dispatch-analysis.json').read_bytes()
        subprocess.run(['python3', '-B', str(P / 'analyze.py'), '--check'], cwd=ROOT, check=True, stdout=subprocess.DEVNULL)
    result = {'status': 'pass', 'helper': desc(__file__), 'perf_analyzer': desc(P / 'analyze.py'), 'dispatch_verifier': desc(P / 'verify_dispatch.py'), 'perf_failure': desc(P / 'first-child-failure.json'), 'dispatch_analysis': desc(P / 'dispatch-analysis.json'), 'descriptor_checks': counts, 'native_source_files_reconstructed': 9355, 'perf_exports_byte_exact': decodes, 'gdb_witnesses_byte_exact': 2, 'perf_failure_byte_exact': True, 'build_or_workload_or_profiler_run': False, 'offline_perf_decoder_run': True, 'deleted_capture_cwd_used': False}
    with args.output.open('x') as stream:
        json.dump(result, stream, indent=2, sort_keys=True)
        stream.write('\n')
    print('0912 native source/perf rejection/GDB witness replay PASS')


if __name__ == '__main__':
    main()
