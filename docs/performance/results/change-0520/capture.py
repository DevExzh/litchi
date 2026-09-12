"""Capture the frozen 0520 native phase baseline, serially in fresh children."""
import datetime
import hashlib
import json
import os
from pathlib import Path
import subprocess
import time

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
BINARY = Path('/home/zhuhe/litchi-goal-0520-target/release/litchi-perf-baseline')


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def source_check():
    for name, digest in json.loads((HERE / 'source-manifest.json').read_text()).items():
        assert sha(REPO / name) == digest, name


def capture(name, command):
    source_check()
    build = json.loads((HERE / 'build.json').read_text())
    assert build['exit_code'] == 0
    assert sha(BINARY) == build['binary_sha256']
    assert not (HERE / (name + '.receipt.json')).exists(), name
    started = datetime.datetime.now(datetime.timezone.utc).isoformat()
    tick = time.monotonic()
    with (HERE / (name + '.stdout')).open('w') as out:
        with (HERE / (name + '.stderr')).open('w') as err:
            result = subprocess.run(command, cwd=REPO, stdout=out, stderr=err)
    source_check()
    assert sha(BINARY) == build['binary_sha256']
    receipt = dict(command=command, start_utc=started,
                   end_utc=datetime.datetime.now(datetime.timezone.utc).isoformat(),
                   seconds=time.monotonic() - tick, exit_code=result.returncode,
                   binary_sha256=sha(BINARY), script_sha256=sha(Path(__file__)),
                   source_manifest_sha256=sha(HERE / 'source-manifest.json'),
                   plan_sha256=sha(HERE / 'plan.json'),
                   environment={k: os.environ.get(k) for k in ['RUSTFLAGS', 'LD_PRELOAD', 'MALLOC_CONF', 'GLIBC_TUNABLES']})
    receipt['artifacts'] = {p.name: sha(p) for p in HERE.glob(name + '.*')
                            if p.is_file() and not p.name.endswith('.receipt.json')}
    (HERE / (name + '.receipt.json')).write_text(json.dumps(receipt, indent=2) + '\n')
    assert result.returncode == 0, (name, result.returncode)
    print(name, 'passed', flush=True)


def main():
    plan = json.loads((HERE / 'plan.json').read_text())
    for repeat in range(1, plan['native_repeats'] + 1):
        shapes = plan['shapes'] if repeat == 1 else list(reversed(plan['shapes']))
        for shape in shapes:
            name = f'native-r{repeat}-{shape}'
            capture(name, ['taskset', '-c', str(plan['cpu']), str(BINARY),
                           '--warmup', str(plan['warmup']), '--samples', str(plan['samples']),
                           '--case', plan['case'], '--xlsx-cell-crud-shape', shape,
                           '--json', str(HERE / (name + '.json'))])


if __name__ == '__main__':
    main()
