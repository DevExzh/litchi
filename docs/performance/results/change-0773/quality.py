"""Root-owned serial post-integration gates; retain failed attempts."""
import hashlib
import json
import os
from pathlib import Path
import subprocess
import time

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
TARGET = Path('/home/zhuhe/code/litchi-target-0773-quality')
OWNERS = ['litchi-core', 'litchi-opc', 'litchi-cfb', 'litchi-doc', 'litchi-docx',
          'litchi-ppt', 'litchi-pptx', 'litchi-xls', 'litchi-xlsx', 'litchi-xlsb']


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def census():
    names = subprocess.check_output(['git', 'ls-files', '-z'], cwd=ROOT).decode().split('\0')
    return {n: sha(ROOT / n) for n in names if n and
            (n.startswith('crates/') or n.startswith('tools/perf-baseline/')) and
            (n.endswith('.rs') or Path(n).name in ('Cargo.toml', 'Cargo.lock'))}


if __name__ == '__main__':
    before = census()
    attempt = 0
    while (P / f'quality-{attempt}').exists():
        attempt += 1
    folder = P / f'quality-{attempt}'
    folder.mkdir()
    (folder / 'source.json').write_text(json.dumps(before, indent=2) + '\n')
    packages = [arg for name in OWNERS for arg in ('-p', name)]
    format_packages = [arg for name in OWNERS[3:] for arg in ('-p', name)]
    locked = ['--offline', '--locked']
    commands = [
        ['cargo', 'fmt', *packages, '--', '--check'],
        ['cargo', 'check', *packages, *locked, '--all-features', '--all-targets'],
        ['cargo', 'test', '-p', 'litchi-core', '-p', 'litchi-opc', '-p', 'litchi-cfb',
         *locked, '--all-features'],
        ['cargo', 'test', *format_packages, *locked, '--all-features', '--test', 'save_durability'],
        ['cargo', 'clippy', *packages, *locked, '--all-features', '--lib', '--', '-D', 'warnings'],
        ['cargo', 'doc', *packages, *locked, '--all-features', '--no-deps'],
        ['cargo', 'test', '--manifest-path', 'tools/perf-baseline/Cargo.toml', *locked, 'save_durability'],
        ['python3', '-B', 'tools/check_crate_boundaries.py'],
    ]
    env = os.environ | {'CARGO_TARGET_DIR': str(TARGET), 'CARGO_BUILD_JOBS': '2',
                       'CARGO_INCREMENTAL': '0', 'CARGO_PROFILE_DEV_DEBUG': '0',
                       'RUSTDOCFLAGS': '-D warnings', 'PYTHONDONTWRITEBYTECODE': '1'}
    rows = []
    for index, command in enumerate(commands):
        log = folder / f'{index:02d}.log'
        start = time.time()
        with log.open('w') as output:
            result = subprocess.run(command, cwd=ROOT, env=env, stdout=output, stderr=subprocess.STDOUT)
        rows.append({'command': command, 'exit': result.returncode, 'started': start,
                     'ended': time.time(), 'log': str(log.relative_to(P)), 'sha256': sha(log)})
        receipt = {'target': str(TARGET), 'source_sha256': sha(folder / 'source.json'),
                   'source_file': str((folder / 'source.json').relative_to(P)), 'rows': rows}
        (folder / 'manifest.json').write_text(json.dumps(receipt, indent=2) + '\n')
        print(f'gate {index + 1}/{len(commands)} exit {result.returncode}', flush=True)
        assert result.returncode == 0, rows[-1]
        assert census() == before, 'source changed during quality gates'
    (P / 'quality.json').write_text(json.dumps(receipt, indent=2) + '\n')
