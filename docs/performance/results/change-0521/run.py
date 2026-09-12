"""Serial, source-bound build and capture driver for the 0521 candidate."""
import argparse
import datetime
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import time

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
SCRATCH = Path('/tmp/litchi-goal-0521')
TARGET = Path('/home/zhuhe/litchi-goal-0521-target')


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def write(path, value):
    path.write_text(json.dumps(value, indent=2) + '\n')


def now():
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def freeze(stage):
    folder = HERE / stage
    folder.mkdir(exist_ok=False)
    names = subprocess.check_output(['git', 'ls-files', '-z', 'crates', 'tools/perf-baseline',
                                     'Cargo.toml', 'Cargo.lock', '.cargo', 'rust-toolchain.toml'], cwd=REPO).split(b'\0')
    paths = {name.decode() for name in names if name}
    paths.update(str(p.relative_to(REPO)) for p in (REPO/'crates/litchi-xlsx/src/cell_values').glob('validation_borrow_tests.rs'))
    write(folder/'source-manifest.json', {name: sha(REPO/name) for name in sorted(paths) if (REPO/name).is_file()})
    (folder/'source.patch').write_bytes(subprocess.check_output(['git', 'diff', '--', 'crates', 'tools/perf-baseline'], cwd=REPO))
    print('Frozen', stage, flush=True)


def check_source(stage):
    for name, digest in json.loads((HERE/stage/'source-manifest.json').read_text()).items():
        assert sha(REPO/name) == digest, name


def run(stage, name, command, binary=None):
    folder = HERE/stage
    receipt_path = folder/(name+'.receipt.json')
    assert not receipt_path.exists(), receipt_path
    check_source(stage)
    before = sha(binary) if binary else None
    start, tick = now(), time.monotonic()
    with (folder/(name+'.stdout')).open('w') as out, (folder/(name+'.stderr')).open('w') as err:
        result = subprocess.run(command, cwd=REPO, stdout=out, stderr=err)
    check_source(stage)
    assert binary is None or sha(binary) == before
    receipt = dict(command=command, start_utc=start, end_utc=now(), seconds=time.monotonic()-tick,
                   exit_code=result.returncode, binary_sha256=before,
                   source_manifest_sha256=sha(folder/'source-manifest.json'),
                   script_sha256=sha(Path(__file__)), plan_sha256=sha(HERE/'plan.json'),
                   environment={key: os.environ.get(key) for key in ['RUSTFLAGS', 'CARGO_ENCODED_RUSTFLAGS', 'LD_PRELOAD', 'MALLOC_CONF', 'GLIBC_TUNABLES']})
    receipt['artifacts'] = {p.name: sha(p) for p in sorted(folder.glob(name+'.*')) if p.is_file() and p != receipt_path}
    write(receipt_path, receipt)
    assert result.returncode == 0, (name, result.returncode)
    print(stage, name, 'passed', flush=True)


def build(stage, kind):
    executable = 'litchi-perf-baseline' + ('-alloc' if kind == 'alloc' else '')
    command = ['cargo', 'build', '--release', '--locked', '--manifest-path', 'tools/perf-baseline/Cargo.toml',
               '--bin', executable, '--target-dir', str(TARGET)]
    if kind == 'alloc':
        command += ['--features', 'allocator-metrics']
    run(stage, 'build-'+kind, command)
    source = TARGET/'release'/executable
    binary = SCRATCH/(stage+'-'+kind)
    shutil.copy2(source, binary)
    assert sha(source) == sha(binary)
    write(HERE/stage/('binary-'+kind+'.json'), dict(path=str(binary), sha256=sha(binary), bytes=binary.stat().st_size,
          build_receipt_sha256=sha(HERE/stage/('build-'+kind+'.receipt.json')),
          source_manifest_sha256=sha(HERE/stage/'source-manifest.json')))


def capture(stage, lane):
    plan = json.loads((HERE/'plan.json').read_text())
    binary = SCRATCH/(stage+('-alloc' if lane == 'alloc' else '-normal'))
    identity = json.loads((HERE/stage/('binary-'+('alloc' if lane=='alloc' else 'normal')+'.json')).read_text())
    assert sha(binary) == identity['sha256']
    jobs = []
    if lane == 'native':
        for repeat in range(1, plan['primary']['repeats']+1):
            shapes = plan['primary']['shapes'][::1 if repeat == 1 else -1]
            for shape in shapes:
                jobs.append((f'native-r{repeat}-primary-{shape}', plan['primary']['case'], shape, plan['primary']['warmup'], plan['primary']['samples']))
        for repeat in range(1, plan['guard_repeats']+1):
            for i, guard in enumerate(plan['guards']):
                for shape in guard['shapes']:
                    jobs.append((f'native-r{repeat}-guard{i}-{shape}', guard['case'], shape, plan['guard_warmup'], plan['guard_samples']))
    else:
        config = plan['allocation' if lane == 'alloc' else 'profile']
        for repeat in range(1, config['repeats']+1):
            for shape in config['shapes']:
                jobs.append((f'{lane}-r{repeat}-{shape}', plan['primary']['case'], shape, config['warmup'], config['samples']))
    for name, case, shape, warmup, samples in jobs:
        command = ['taskset', '-c', str(plan['cpu'])]
        if lane == 'native':
            command += ['/usr/bin/time', '-f', '{"max_rss_kib":%M,"elapsed_seconds":%e,"user_seconds":%U,"system_seconds":%S}',
                        '-o', str(HERE/stage/(name+'.rss.json'))]
        elif lane == 'profile':
            owner = plan['profile']['owner']
            command += ['valgrind', '--tool=callgrind', '--collect-atstart=no', '--toggle-collect='+owner,
                        '--zero-before='+owner, '--dump-after='+owner,
                        '--callgrind-out-file='+str(HERE/stage/(name+'.callgrind'))]
        command += [str(binary), '--warmup', str(warmup), '--samples', str(samples), '--case', case,
                    '--xlsx-cell-crud-shape', shape, '--json', str(HERE/stage/(name+'.json'))]
        run(stage, name, command, binary)


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('stage', choices=['baseline', 'candidate'])
    parser.add_argument('action', choices=['freeze', 'build-normal', 'build-alloc', 'native', 'profile', 'alloc'])
    args = parser.parse_args()
    if args.action == 'freeze':
        freeze(args.stage)
    elif args.action.startswith('build-'):
        build(args.stage, args.action[6:])
    else:
        capture(args.stage, args.action)
