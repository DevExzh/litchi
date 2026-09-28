"""Root-only serial command recorder for preservation validation and export."""
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import time

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
PLAN = json.loads((P / 'plan.json').read_text())


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def write(path, value):
    assert not path.exists(), path
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + '\n')


def source():
    names = subprocess.check_output(['git', 'ls-files', '-z', '--cached', '--others', '--exclude-standard', 'crates', 'Cargo.toml', 'Cargo.lock', 'rustfmt.toml', 'tools/perf-baseline'], cwd=ROOT).decode().split('\0')
    return {name: sha(ROOT / name) for name in names if name and (ROOT / name).is_file()}


def guard():
    assert subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=ROOT, text=True).strip() == PLAN['base']
    for name, wanted in PLAN['architecture'].items():
        assert sha(ROOT / name) == wanted, name
    for name, wanted in PLAN['unrelated'].items():
        assert sha(ROOT / name) == wanted, name
    assert sha(ROOT / 'Cargo.lock') == PLAN['root_lock_sha256']
    assert sha(ROOT / 'tools/perf-baseline/Cargo.lock') == PLAN['tool_lock_sha256']
    for name, expected in PLAN['admission']['real_inputs'].items():
        path = ROOT / name
        assert sha(path) == expected['sha256'] and path.stat().st_size == expected['bytes']


def main():
    assert len(sys.argv) >= 4, 'run.py LABEL quality|release COMMAND...'
    label, mode, *command = sys.argv[1:]
    assert label.replace('-', '').replace('_', '').isalnum()
    assert mode in ('quality', 'release')
    guard()
    for name in ('RUSTFLAGS', 'CARGO_ENCODED_RUSTFLAGS', 'RUSTC_WRAPPER', 'RUSTC_WORKSPACE_WRAPPER'):
        assert not os.environ.get(name), name
    before = source()
    source_encoded = (json.dumps(before, indent=2, sort_keys=True) + '\n').encode()
    source_sha = hashlib.sha256(source_encoded).hexdigest()
    snapshots = P / 'sources'
    snapshots.mkdir(exist_ok=True)
    snapshot = snapshots / (source_sha + '.json')
    if snapshot.exists():
        assert snapshot.read_bytes() == source_encoded
    else:
        snapshot.write_bytes(source_encoded)
    out = P / 'commands' / label
    out.mkdir(parents=True, exist_ok=False)
    driver_sha, plan_sha = sha(Path(__file__)), sha(P / 'plan.json')
    env_values = {'CARGO_TARGET_DIR': PLAN['target'], 'CARGO_BUILD_JOBS': '2', 'CARGO_INCREMENTAL': '0'}
    if mode == 'release':
        env_values.update({'CARGO_PROFILE_RELEASE_OPT_LEVEL': '3', 'CARGO_PROFILE_RELEASE_DEBUG': '1', 'CARGO_PROFILE_RELEASE_LTO': 'thin', 'CARGO_PROFILE_RELEASE_CODEGEN_UNITS': '1', 'CARGO_PROFILE_RELEASE_INCREMENTAL': 'false', 'CARGO_PROFILE_RELEASE_PANIC': 'unwind'})
    receipt = {'schema': 'litchi.performance.0818.command.v1', 'label': label, 'command': command, 'cwd': str(ROOT), 'environment': env_values, 'source': {'path': str(snapshot.relative_to(P)), 'sha256': source_sha}, 'driver_sha256': driver_sha, 'plan_sha256': plan_sha, 'started': time.time()}
    write(out / 'started.json', receipt)
    with (out / 'output.log').open('w') as stream:
        result = subprocess.run(command, cwd=ROOT, env=os.environ | env_values, stdout=stream, stderr=subprocess.STDOUT)
    receipt.update(ended=time.time(), exit_code=result.returncode, log_sha256=sha(out / 'output.log'))
    write(out / 'result.json', receipt)
    guard()
    assert source() == before, 'source changed during command'
    assert sha(Path(__file__)) == driver_sha and sha(P / 'plan.json') == plan_sha
    print(label, 'terminal exit', result.returncode, flush=True)
    return result.returncode


if __name__ == '__main__':
    raise SystemExit(main())
