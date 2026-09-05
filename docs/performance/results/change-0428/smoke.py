#!/usr/bin/env python3
"""Exercise the real diagnostic CLI before the release protocol is frozen."""
import argparse
import hashlib
import json
from pathlib import Path
import subprocess
import sys
import tempfile

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--directory', default='smoke')
    parser.add_argument('--binary', type=Path,
                        default=REPO / 'tools/perf-baseline/target/debug/litchi-perf-baseline')
    args = parser.parse_args()
    binary = args.binary.resolve()
    revision = subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=REPO).decode().strip()
    assert Path(args.directory).name == args.directory and args.directory not in {'.', '..'}
    directory = ROOT / args.directory
    directory.mkdir(exist_ok=False)
    commands = []

    def run(argv, expected_success):
        result = subprocess.run([str(arg) for arg in argv], cwd=REPO, capture_output=True, text=True)
        commands.append({'argv': [str(arg) for arg in argv], 'exit_code': result.returncode,
                         'stdout': result.stdout, 'stderr': result.stderr,
                         'expected_success': expected_success})
        (directory / 'commands-in-progress.json').write_text(json.dumps(commands, indent=2) + '\n')
        assert (result.returncode == 0) == expected_success, commands[-1]
        print(json.dumps(commands[-1]), flush=True)

    protocol = json.loads((ROOT / 'protocol.json').read_text())
    for lane in protocol['lanes']:
        report = directory / (lane['scenario'] + '-' + lane['corpus'] + '.json')
        argv = [binary, 'cache-retention', '--scenario', lane['scenario'], '--corpus', lane['corpus'],
                '--samples', '1', '--warmup', '0', '--source-revision', revision, '--output', report]
        run(argv, True)
        run([sys.executable, '-B', ROOT / 'verify-report.py', report], True)
        run([sys.executable, '-B', ROOT / 'probe-report.py', report], True)
    with tempfile.TemporaryDirectory(prefix='litchi-0428-cli-') as temporary:
        path = Path(temporary) / 'output.json'
        valid = [binary, 'cache-retention', '--scenario', 'lifecycle', '--corpus', 'plain', '--samples', '1',
                 '--warmup', '0', '--source-revision', revision, '--output', path]
        for flag in ['--scenario', '--corpus', '--samples', '--warmup', '--source-revision', '--output']:
            index = valid.index(flag)
            run(valid[:index] + valid[index + 2:], False)
            assert not path.exists()
        for flag, value in [('--samples', '0'), ('--samples', '1001'), ('--warmup', '1001'),
                            ('--source-revision', 'invalid'), ('--scenario', 'unknown'), ('--corpus', 'unknown')]:
            argv = valid.copy()
            argv[argv.index(flag) + 1] = value
            run(argv, False)
            assert not path.exists()
        run(valid + ['--scenario', 'lifecycle'], False)
        assert not path.exists()
        path.write_bytes(b'preserve existing output\n')
        run(valid, False)
        assert path.read_bytes() == b'preserve existing output\n'
    (directory / 'commands.json').write_text(json.dumps({
        'classification': 'CLI correctness validation; not formal measurements',
        'revision': revision, 'binary_sha256': hashlib.sha256(binary.read_bytes()).hexdigest(),
        'commands': commands, 'status': 'pass'}, indent=2) + '\n')


if __name__ == '__main__':
    main()
