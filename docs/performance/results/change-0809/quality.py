"""Root-owned, serial PPTX quality gates for the test-only baseline repair."""
import hashlib
import json
import os
from pathlib import Path
import re
import subprocess
import time

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
TARGET = ROOT.parent / 'litchi-target-0809'

def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()

def write(path, value):
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + '\n')

def source():
    names = subprocess.check_output(['git', 'ls-files', '-z', '--', 'crates', 'Cargo.toml', 'clippy.toml', '.cargo/config.toml', 'rust-toolchain.toml'], cwd=ROOT).decode().split('\0')
    return {name: sha(ROOT / name) for name in names if name}

assert not TARGET.exists()
assert not (P / 'checks.json').exists()
origin = json.loads((P / 'origin.json').read_text())
assert subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=ROOT, text=True).strip() == origin['base']
before = json.loads((P / origin['baseline_source_reference']).read_text())['files']
current = source()
assert len(current) == 9196
assert {name for name in before.keys() | current.keys() if before.get(name) != current.get(name)} == set(origin['scope'])
write(P / 'source.json', current)
commands = [
    ['cargo', 'fmt', '-p', 'litchi-pptx', '--', '--check'],
    ['cargo', 'check', '--offline', '--locked', '-p', 'litchi-pptx', '--all-features', '--all-targets'],
    ['cargo', 'test', '--offline', '--locked', '-p', 'litchi-pptx', '--all-features', '--', '--test-threads=2'],
    ['cargo', 'clippy', '--offline', '--locked', '-p', 'litchi-pptx', '--all-features', '--all-targets', '--', '-D', 'warnings'],
    ['cargo', 'doc', '--offline', '--locked', '-p', 'litchi-pptx', '--all-features', '--no-deps'],
    ['python3', '-B', 'tools/check_crate_boundaries.py'],
]
env_values = {'CARGO_TARGET_DIR': str(TARGET), 'CARGO_BUILD_JOBS': '2', 'CARGO_INCREMENTAL': '0', 'CARGO_PROFILE_DEV_DEBUG': '0', 'RUSTDOCFLAGS': '-D warnings', 'PYTHONDONTWRITEBYTECODE': '1'}
write(P / 'plan.json', {'commands': commands, 'environment': env_values, 'driver_sha256': sha(__file__), 'source_sha256': sha(P / 'source.json')})
rows = []
for index, command in enumerate(commands):
    log = P / f'{index + 1:02}.log'
    started = time.time()
    with log.open('w') as output:
        result = subprocess.run(command, cwd=ROOT, env=os.environ | env_values, stdout=output, stderr=subprocess.STDOUT)
    rows.append({'gate': index + 1, 'command': command, 'started': started, 'ended': time.time(), 'exit_code': result.returncode, 'log': {'path': log.name, 'bytes': log.stat().st_size, 'sha256': sha(log)}})
    write(P / 'checks.json', rows)
    assert result.returncode == 0, log
    assert source() == current
    print(f'0809 gate {index + 1} PASS', flush=True)
counts = [tuple(map(int, values)) for values in re.findall(r'test result: ok\. (\d+) passed; (\d+) failed; (\d+) ignored; (\d+) measured;', (P / '03.log').read_text())]
assert counts and all(row[1] == 0 for row in counts)
write(P / 'complete.json', {'schema': 'litchi.performance.0809.quality.v1', 'gates_passed': 6, 'test_groups': len(counts), 'tests': dict(zip(('passed', 'failed', 'ignored', 'measured'), (sum(row[i] for row in counts) for i in range(4)))), 'checks_sha256': sha(P / 'checks.json'), 'source_sha256': sha(P / 'source.json'), 'plan_sha256': sha(P / 'plan.json')})
