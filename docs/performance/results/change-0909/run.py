"""Create-only custody runner for the untimed 0909 owning-input qualification."""
import argparse
import datetime
import hashlib
import json
import os
from pathlib import Path
import subprocess
import shutil
import tempfile

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
SOURCE = Path('/tmp/litchi-0909-src')
TARGET = ROOT / 'target/litchi-0909-build'
PARENT = '57de5495f899a43574110f8c572823a7d9907e5f'
OVERLAYS = ['tools/perf-baseline/src/lib.rs', 'tools/perf-baseline/src/pptx_owned_control.rs', 'tools/perf-baseline/examples/pptx_owned_control.rs']


def descriptor(path):
    data = path.read_bytes()
    return {'path': str(path), 'bytes': len(data), 'sha256': hashlib.sha256(data).hexdigest()}


def write_new(path, obj):
    with path.open('x') as stream:
        json.dump(obj, stream, indent=2)
        stream.write('\n')


def now():
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('mode', choices=['freeze', 'build', 'capture'])
    parser.add_argument('--attempt', default='v1')
    args = parser.parse_args()
    assert args.attempt.replace('-', '').isalnum()
    assert (SOURCE / '.owned-0909').read_text() == PARENT + '\n'
    assert (TARGET / '.owned-0909').read_text() == PARENT + '\n'
    if args.mode == 'freeze':
        with tempfile.TemporaryDirectory(prefix='litchi-0909-overlay-index-') as directory:
            env = dict(os.environ, GIT_INDEX_FILE=str(Path(directory) / 'index'))
            subprocess.run(['git', 'read-tree', PARENT], cwd=ROOT, env=env, check=True)
            subprocess.run(['git', 'add', '--', *OVERLAYS], cwd=ROOT, env=env, check=True)
            patch = subprocess.check_output(['git', 'diff', '--cached', '--binary', PARENT, '--', *OVERLAYS], cwd=ROOT, env=env)
        patch_path = P / f'source-overlay-{args.attempt}.patch'
        with patch_path.open('xb') as stream:
            stream.write(patch)
        for name in OVERLAYS:
            dst = SOURCE / name
            dst.parent.mkdir(parents=True, exist_ok=True)
            dst.write_bytes((ROOT / name).read_bytes())
        inventory = []
        for f in sorted(SOURCE.rglob('*')):
            if f.is_file():
                d = descriptor(f)
                d['path'] = str(f.relative_to(SOURCE))
                inventory.append(d)
        write_new(P / f'source-{args.attempt}.json', {'parent': PARENT, 'base': descriptor(P / 'source-base.json'), 'overlay_patch': descriptor(patch_path), 'overlays': [descriptor(ROOT / name) for name in OVERLAYS], 'files': inventory})
        print('Frozen source', len(inventory), 'files')
        return
    manifest = P / f'source-{args.attempt}.json'
    source = json.loads(manifest.read_text())
    for row in source['files']:
        f = SOURCE / row['path']
        assert f.stat().st_size == row['bytes'] and descriptor(f)['sha256'] == row['sha256'], f
    folder = P / f'{args.mode}-{args.attempt}'
    folder.mkdir()
    env = dict(os.environ)
    env.update({'CARGO_TARGET_DIR': str(TARGET), 'CARGO_BUILD_JOBS': '2', 'CARGO_INCREMENTAL': '0', 'CARGO_PROFILE_RELEASE_LTO': 'false', 'CARGO_PROFILE_RELEASE_CODEGEN_UNITS': '16', 'CARGO_PROFILE_RELEASE_DEBUG': '0', 'LC_ALL': 'C', 'LANG': 'C', 'TZ': 'UTC'})
    env.pop('RUSTFLAGS', None)
    env.pop('CARGO_ENCODED_RUSTFLAGS', None)
    binary = TARGET / 'release/examples/pptx_owned_control'
    if args.mode == 'build':
        argv = ['cargo', 'build', '--manifest-path', str(SOURCE / 'tools/perf-baseline/Cargo.toml'), '--locked', '--offline', '--release', '--example', 'pptx_owned_control', '-j', '2']
    else:
        argv = [str(binary), str(folder / 'artifacts')]
    start = {'argv': argv, 'cwd': str(SOURCE), 'environment_overrides': {k: env[k] for k in ('CARGO_TARGET_DIR', 'CARGO_BUILD_JOBS', 'CARGO_INCREMENTAL', 'CARGO_PROFILE_RELEASE_LTO', 'CARGO_PROFILE_RELEASE_CODEGEN_UNITS', 'CARGO_PROFILE_RELEASE_DEBUG', 'LC_ALL', 'LANG', 'TZ')}, 'removed_environment_keys': ['RUSTFLAGS', 'CARGO_ENCODED_RUSTFLAGS'], 'start_utc': now(), 'source': descriptor(manifest), 'helper': descriptor(Path(__file__)), 'scope': 'Untimed correctness qualification; this build is not the 0907 timing binary.'}
    start['toolchain'] = {name: {'launcher': shutil.which(name), 'version': subprocess.check_output([name, '--version'], cwd=SOURCE, env=env, text=True).strip()} for name in ('cargo', 'rustc')}
    if args.mode == 'capture':
        start['binary'] = descriptor(binary)
    write_new(folder / 'start.json', start)
    with (folder / 'stdout.txt').open('x') as out, (folder / 'stderr.txt').open('x') as err:
        result = subprocess.run(argv, cwd=SOURCE, env=env, stdout=out, stderr=err)
    receipt = dict(start, end_utc=now(), exit_code=result.returncode)
    receipt['logs'] = [descriptor(folder / 'stdout.txt'), descriptor(folder / 'stderr.txt')]
    if result.returncode == 0 and args.mode == 'build':
        receipt['binary'] = descriptor(binary)
    if args.mode == 'capture':
        receipt['artifacts'] = [descriptor(f) for f in sorted((folder / 'artifacts').rglob('*')) if f.is_file()]
    write_new(folder / 'command.json', receipt)
    print(args.mode, args.attempt, 'exit', result.returncode)
    raise SystemExit(result.returncode)


if __name__ == '__main__':
    main()
