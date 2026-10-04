"""Scoped formatting and warning-denied Clippy for the qualification example."""
import argparse
import datetime
import json
import os
from pathlib import Path
import subprocess

P = Path(__file__).resolve().parent


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--attempt', required=True)
    args = parser.parse_args()
    build = json.loads((P / f'build-{args.attempt}/command.json').read_text())
    assert build['exit_code'] == 0
    source = Path(build['cwd'])
    folder = P / f'quality-{args.attempt}'
    folder.mkdir()
    env = dict(os.environ, **build['environment_overrides'])
    for key in build['removed_environment_keys']:
        env.pop(key, None)
    commands = [
        ['rustfmt', '--check', '--edition', '2024', '--config', 'skip_children=true', *[str(source / name) for name in ['tools/perf-baseline/src/lib.rs', 'tools/perf-baseline/src/pptx_owned_control.rs', 'tools/perf-baseline/examples/pptx_owned_control.rs']]],
        ['cargo', 'clippy', '--manifest-path', str(source / 'tools/perf-baseline/Cargo.toml'), '--locked', '--offline', '--release', '--example', 'pptx_owned_control', '-j', '2', '--', '-D', 'warnings'],
    ]
    rows = []
    for index, argv in enumerate(commands):
        start = datetime.datetime.now(datetime.timezone.utc).isoformat()
        with (folder / f'{index}.stdout.txt').open('x') as out, (folder / f'{index}.stderr.txt').open('x') as err:
            result = subprocess.run(argv, cwd=source, env=env, stdout=out, stderr=err)
        rows.append({'argv': argv, 'start_utc': start, 'end_utc': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'exit_code': result.returncode})
        if result.returncode:
            break
    receipt = {'status': 'pass' if all(r['exit_code'] == 0 for r in rows) and len(rows) == len(commands) else 'fail', 'build_receipt': str(P / f'build-{args.attempt}/command.json'), 'source': build['source'], 'environment_overrides': build['environment_overrides'], 'commands': rows, 'scope': 'Changed harness lib/example; no production implementation changed; no full workspace test claim.'}
    with (folder / 'receipt.json').open('x') as stream:
        json.dump(receipt, stream, indent=2)
        stream.write('\n')
    print('0909 scoped quality', receipt['status'])
    raise SystemExit(0 if receipt['status'] == 'pass' else 1)


if __name__ == '__main__':
    main()
