"""Root-only repair verification; preserve the failed frozen experiment."""
from pathlib import Path
import json
import os
import subprocess
import sys
import time

P = Path(__file__).resolve().parent
sys.path.insert(0, str(P.parent))
import custody as c

ORIGIN = c.read(P / 'origin.json')
NAME = ORIGIN['tool_allowlist'][0]
BASELINE = c.read(P.parent / 'quality-0/source.json')
FROZEN = c.read(P.parent / 'quality-0/frozen-inputs.json')
MANIFEST = str(c.TOOL / 'Cargo.toml')
PREFIX = ['--offline', '--locked', '--manifest-path', MANIFEST]
COMMANDS = {
    'focused': ['cargo', 'test', *PREFIX, '--all-features', '--bin', 'docx_replayable_tail_append', '--', '--test-threads=2'],
    'fmt': ['cargo', 'fmt', '--manifest-path', MANIFEST, '--', '--check'],
    'check': ['cargo', 'check', *PREFIX, '--all-features', '--all-targets'],
    'tests': ['cargo', 'test', *PREFIX, '--all-features', '--', '--test-threads=2'],
    'clippy': ['cargo', 'clippy', *PREFIX, '--all-features', '--all-targets', '--', '-D', 'warnings'],
    'rustdoc': ['cargo', 'doc', *PREFIX, '--all-features', '--no-deps'],
    'boundaries': ['python3', '-B', 'tools/check_crate_boundaries.py'],
}


def current():
    production, tool = c.source(), c.tool_source()
    assert production == BASELINE['production']
    assert production['revision'] == ORIGIN['base']
    changed = {name for name in tool if tool[name] != BASELINE['tool'][name]}
    assert changed == {NAME}, changed
    assert set(tool) == set(BASELINE['tool'])
    before = (P / 'before-counting_allocator.rs').read_bytes()
    after = (c.ROOT / NAME).read_bytes()
    assert c.sha(P / 'before-counting_allocator.rs') == ORIGIN['before_sha256'] == BASELINE['tool'][NAME]
    assert before.split(b'#[cfg(test)]', 1)[0] == after.split(b'#[cfg(test)]', 1)[0]
    assert c.packet_hashes() == FROZEN['packet']
    assert c.driver_hashes() == FROZEN['drivers']
    assert c.assert_root_inputs() == FROZEN['root_inputs']
    assert c.lock_identity() == FROZEN['locks']
    assert c.architecture_hashes() == FROZEN['architecture']
    assert c.assert_corpus_inputs() == FROZEN['corpus']
    assert c.assert_unrelated() == FROZEN['unrelated']
    assert c.assert_host() == FROZEN['host']
    return {'production': production, 'tool': tool}


def main():
    assert len(sys.argv) == 2
    action = sys.argv[1]
    source = current()
    source_path = P / 'source.json'
    if action == 'freeze':
        assert not source_path.exists()
        c.write(source_path, source)
        c.write(P / 'inputs.json', {'schema': 'litchi.performance.0820.repair-inputs.v1',
                                  'runner': c.artifact(Path(__file__)),
                                  'origin': c.artifact(P / 'origin.json'),
                                  'before': c.artifact(P / 'before-counting_allocator.rs'),
                                  'source': c.artifact(source_path),
                                  'original_frozen_inputs': c.artifact(P.parent / 'quality-0/frozen-inputs.json')})
        print('0820 repair source frozen', flush=True)
        return
    assert action in COMMANDS
    assert c.read(source_path) == source
    frozen = c.read(P / 'inputs.json')
    for key in ('runner', 'origin', 'before', 'source', 'original_frozen_inputs'):
        assert c.artifact(frozen[key]['path']) == frozen[key]
    out = P / 'commands' / action
    assert not out.exists(), f'refusing overwrite {out}'
    out.mkdir(parents=True)
    for name in ('RUSTUP_TOOLCHAIN', 'RUSTFLAGS', 'CARGO_ENCODED_RUSTFLAGS', 'CARGO_BUILD_RUSTFLAGS',
                 'CARGO_TARGET_X86_64_UNKNOWN_LINUX_GNU_RUSTFLAGS', 'RUSTC_WRAPPER', 'RUSTC_WORKSPACE_WRAPPER'):
        assert not os.environ.get(name), name
    environment = {'CARGO_TARGET_DIR': ORIGIN['target'], 'CARGO_BUILD_JOBS': '2',
                   'CARGO_INCREMENTAL': '0', 'CARGO_PROFILE_DEV_DEBUG': '0',
                   'RUSTDOCFLAGS': '-D warnings', 'PYTHONDONTWRITEBYTECODE': '1'}
    command = COMMANDS[action]
    started = time.time()
    start = {'schema': 'litchi.performance.0820.repair-command.v1', 'name': action,
             'command': command, 'started': started, 'environment': environment,
             'inputs': c.artifact(P / 'inputs.json'), 'source': c.artifact(source_path)}
    c.write(out / 'started.json', start)
    log = out / 'output.log'
    with log.open('w') as stream:
        result = subprocess.run(command, cwd=c.ROOT, env=os.environ | environment,
                                stdout=stream, stderr=subprocess.STDOUT)
    receipt = start | {'ended': time.time(), 'exit_code': result.returncode, 'log': c.artifact(log)}
    c.write(out / 'result.json', receipt)
    assert current() == source
    assert result.returncode == 0, f'{action} failed; retained {log}'
    print('0820 repair', action, 'PASS', flush=True)


if __name__ == '__main__':
    main()
