"""Create-only source/build/native capture custody for the 0910 route control."""
import argparse
import datetime
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import tempfile

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
SOURCE = Path('/tmp/litchi-0910-src')
TARGET = ROOT / 'target/litchi-0910-build'
PARENT = '468fe2931d6532c32f4c6608590816b8eaa88e31'
OVERLAYS = ['tools/perf-baseline/src/lib.rs', 'tools/perf-baseline/src/pptx_owned_benchmark.rs', 'tools/perf-baseline/examples/pptx_owned_benchmark.rs']


def descriptor(path):
    data = path.read_bytes()
    return {'path': str(path.relative_to(ROOT)) if path.is_relative_to(ROOT) else str(path), 'bytes': len(data), 'sha256': hashlib.sha256(data).hexdigest()}


def write_new(path, obj):
    with path.open('x') as stream:
        json.dump(obj, stream, indent=2)
        stream.write('\n')


def now():
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('mode', choices=['freeze', 'build', 'capture', 'quality'])
    parser.add_argument('--attempt', default='v1')
    args = parser.parse_args()
    assert args.attempt.replace('-', '').isalnum()
    for d in (SOURCE, TARGET):
        assert (d / '.owned-0910').read_text() == PARENT + '\n'
    if args.mode == 'freeze':
        with tempfile.TemporaryDirectory(prefix='litchi-0910-overlay-index-') as directory:
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
        print('Frozen source', len(inventory), 'files', flush=True)
        return
    manifest = P / f'source-{args.attempt}.json'
    for row in json.loads(manifest.read_text())['files']:
        f = SOURCE / row['path']
        assert f.stat().st_size == row['bytes'] and descriptor(f)['sha256'] == row['sha256'], f
    env = dict(os.environ)
    env.update({'CARGO_TARGET_DIR': str(TARGET), 'CARGO_BUILD_JOBS': '2', 'CARGO_INCREMENTAL': '0', 'CARGO_PROFILE_RELEASE_LTO': 'false', 'CARGO_PROFILE_RELEASE_CODEGEN_UNITS': '16', 'CARGO_PROFILE_RELEASE_DEBUG': '0', 'LC_ALL': 'C', 'LANG': 'C', 'TZ': 'UTC'})
    for key in ('RUSTFLAGS', 'CARGO_ENCODED_RUSTFLAGS'):
        env.pop(key, None)
    binary = TARGET / 'release/examples/pptx_owned_benchmark'
    common = {'cwd': str(SOURCE), 'environment_overrides': {k: env[k] for k in ('CARGO_TARGET_DIR', 'CARGO_BUILD_JOBS', 'CARGO_INCREMENTAL', 'CARGO_PROFILE_RELEASE_LTO', 'CARGO_PROFILE_RELEASE_CODEGEN_UNITS', 'CARGO_PROFILE_RELEASE_DEBUG', 'LC_ALL', 'LANG', 'TZ')}, 'removed_environment_keys': ['RUSTFLAGS', 'CARGO_ENCODED_RUSTFLAGS'], 'source': descriptor(manifest), 'helper': descriptor(Path(__file__)), 'toolchain': {n: {'launcher': shutil.which(n), 'version': subprocess.check_output([n, '--version'], cwd=SOURCE, env=env, text=True).strip()} for n in ('cargo', 'rustc')}}
    def execute(folder, argv, extra=None):
        folder.mkdir(parents=True)
        start = dict(common, argv=argv, start_utc=now(), **(extra or {}))
        write_new(folder / 'start.json', start)
        with (folder / 'stdout.txt').open('x') as out, (folder / 'stderr.txt').open('x') as err:
            result = subprocess.run(argv, cwd=SOURCE, env=env, stdout=out, stderr=err)
        receipt = dict(start, end_utc=now(), exit_code=result.returncode)
        receipt['logs'] = [descriptor(folder / name) for name in ('stdout.txt', 'stderr.txt')]
        if (folder / 'artifacts').is_dir():
            receipt['artifacts'] = [descriptor(f) for f in sorted((folder / 'artifacts').rglob('*')) if f.is_file()]
        if (folder / 'process.txt').exists():
            receipt['process_metrics'] = descriptor(folder / 'process.txt')
        if args.mode == 'build' and result.returncode == 0:
            receipt['binary'] = descriptor(binary)
        write_new(folder / 'command.json', receipt)
        print(folder.name, 'exit', result.returncode, flush=True)
        if result.returncode:
            raise SystemExit(result.returncode)
        return receipt
    if args.mode == 'build':
        execute(P / f'build-{args.attempt}', ['cargo', 'build', '--manifest-path', str(SOURCE / 'tools/perf-baseline/Cargo.toml'), '--locked', '--offline', '--release', '--example', 'pptx_owned_benchmark', '-j', '2'])
    elif args.mode == 'quality':
        folder = P / f'quality-{args.attempt}'
        execute(folder / 'fmt', ['rustfmt', '--check', '--edition', '2024', '--config', 'skip_children=true', '--config-path', str(SOURCE / 'rustfmt.toml'), *[str(SOURCE / n) for n in OVERLAYS]])
        execute(folder / 'clippy', ['cargo', 'clippy', '--manifest-path', str(SOURCE / 'tools/perf-baseline/Cargo.toml'), '--locked', '--offline', '--release', '--example', 'pptx_owned_benchmark', '-j', '2', '--', '-D', 'warnings'])
    else:
        plan = json.loads((P / 'capture-plan.json').read_text())
        assert plan['status'] == 'frozen_before_capture'
        out = P / 'capture'
        out.mkdir()
        for leg in plan['legs']:
            folder = out / leg['label']
            argv = ['/usr/bin/time', '-f', 'max_rss_kib=%M\nuser_seconds=%U\nsystem_seconds=%S\nelapsed_seconds=%e\nmajor_faults=%F\nminor_faults=%R\nvoluntary_context_switches=%w\ninvoluntary_context_switches=%c', '-o', str(folder / 'process.txt'), 'taskset', '-c', str(plan['cpu']), str(binary), leg['route'], str(plan['warmup']), str(plan['samples']), str(folder / 'artifacts')]
            execute(folder, argv, {'binary': descriptor(binary), 'plan': descriptor(P / 'capture-plan.json'), 'leg': leg})


if __name__ == '__main__':
    main()
