"""Offline replay of 0809 source scope, exact test repair, and six quality gates."""
import hashlib
import json
from pathlib import Path
import re
import subprocess
import sys

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
TARGET = ROOT.parent / 'litchi-target-0809'

def read(path):
    return json.loads(path.read_text())

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def validate():
    origin = read(P / 'origin.json')
    for name in ('architecture', 'baseline_source'):
        assert sha(P / origin[name + '_reference']) == origin[name + '_sha256']
    for name, digest in read(P / origin['architecture_reference']).items():
        assert sha(ROOT / name) == digest, name
    for name, digest in origin['unrelated'].items():
        assert sha(ROOT / name) == digest, name
    additional = read(P / 'additional-inputs.json')
    assert {row['repository_path'] for row in additional['inputs']} == {'Cargo.lock', 'rustfmt.toml'}
    for row in additional['inputs']:
        original_input = ROOT / row['repository_path']
        retained_input = P / row['retained_path']
        assert original_input.read_bytes() == retained_input.read_bytes()
        assert sha(retained_input) == row['sha256'] and retained_input.stat().st_size == row['bytes']
        assert row['mtime_ns'] / 1e9 < additional['first_gate_started']
    assert subprocess.check_output(['git', 'show', origin['base'] + ':rustfmt.toml'], cwd=ROOT) == (P / 'inputs/rustfmt.toml').read_bytes()
    source = read(P / 'source.json')
    before = read(P / origin['baseline_source_reference'])['files']
    names = subprocess.check_output(['git', 'ls-files', '-z', '--', 'crates', 'Cargo.toml', 'clippy.toml', '.cargo/config.toml', 'rust-toolchain.toml'], cwd=ROOT).decode().split('\0')
    assert set(source) == {name for name in names if name} and len(source) == 9196
    assert {name for name in before.keys() | source.keys() if before.get(name) != source.get(name)} == set(origin['scope'])
    for name, digest in source.items():
        assert sha(ROOT / name) == digest, name
    name = origin['scope'][0]
    original = subprocess.check_output(['git', 'show', origin['base'] + ':' + name], cwd=ROOT).decode()
    assert hashlib.sha256(original.encode()).hexdigest() == origin['before_sha256']
    expected = original
    for indent, message in ((12, 'exact invalid proofs must refuse the notes graph'), (8, 'the 16 MiB slide root limit must refuse capture'), (8, 'the 64 MiB part limit must refuse capture')):
        old = '.err()\n' + ' ' * indent + f'.expect("{message}")'
        assert expected.count(old) == 1
        expected = expected.replace(old, f'.expect_err("{message}")')
    assert (ROOT / name).read_text() == expected
    plan = read(P / 'plan.json')
    assert plan['driver_sha256'] == sha(P / 'quality.py') and plan['source_sha256'] == sha(P / 'source.json')
    commands = [
        ['cargo', 'fmt', '-p', 'litchi-pptx', '--', '--check'],
        ['cargo', 'check', '--offline', '--locked', '-p', 'litchi-pptx', '--all-features', '--all-targets'],
        ['cargo', 'test', '--offline', '--locked', '-p', 'litchi-pptx', '--all-features', '--', '--test-threads=2'],
        ['cargo', 'clippy', '--offline', '--locked', '-p', 'litchi-pptx', '--all-features', '--all-targets', '--', '-D', 'warnings'],
        ['cargo', 'doc', '--offline', '--locked', '-p', 'litchi-pptx', '--all-features', '--no-deps'],
        ['python3', '-B', 'tools/check_crate_boundaries.py'],
    ]
    assert plan['commands'] == commands
    assert plan['environment']['RUSTDOCFLAGS'] == '-D warnings'
    assert plan['environment']['CARGO_TARGET_DIR'] == str(TARGET)
    rows = read(P / 'checks.json')
    assert len(rows) == 6
    previous_end = 0
    for index, row in enumerate(rows):
        assert row['gate'] == index + 1 and row['command'] == commands[index] and row['exit_code'] == 0
        assert previous_end <= row['started'] <= row['ended']
        previous_end = row['ended']
        log = P / row['log']['path']
        assert log.name == f'{index + 1:02}.log' and log.stat().st_size == row['log']['bytes'] and sha(log) == row['log']['sha256']
    counts = [tuple(map(int, values)) for values in re.findall(r'test result: ok\. (\d+) passed; (\d+) failed; (\d+) ignored; (\d+) measured;', (P / '03.log').read_text())]
    totals = dict(zip(('passed', 'failed', 'ignored', 'measured'), (sum(row[i] for row in counts) for i in range(4))))
    complete = read(P / 'complete.json')
    assert complete['gates_passed'] == 6 and complete['tests'] == totals and complete['test_groups'] == len(counts)
    assert totals['failed'] == 0
    for name in ('checks', 'source', 'plan'):
        assert complete[name + '_sha256'] == sha(P / (name + '.json'))
    if '--require-cleanup' in sys.argv:
        cleanup = read(P / 'cleanup.json')
        assert cleanup['target'] == str(TARGET) and cleanup['removed'] is True and not TARGET.exists()
        assert cleanup['source_sha256'] == sha(P / 'source.json')
    print('0809 validation PASS: exact three test repairs, six gates,', totals)

if __name__ == '__main__':
    validate()
