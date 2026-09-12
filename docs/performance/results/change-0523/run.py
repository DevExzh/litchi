"""Serial source-bound current CFB/OLE2 build and measurement driver."""
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
SCRATCH = Path('/tmp/litchi-goal-0523')
TARGET = Path('/home/zhuhe/litchi-goal-0523-target')
FOLDER = HERE / 'baseline'


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def write(path, value):
    path.write_text(json.dumps(value, indent=2) + '\n')


def now():
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def freeze():
    FOLDER.mkdir(exist_ok=False)
    names = subprocess.check_output(['git', 'ls-files', '-z', 'crates',
        'tools/perf-baseline', 'Cargo.toml', 'Cargo.lock', '.cargo',
        'rust-toolchain.toml'], cwd=REPO).split(b'\0')
    paths = sorted(name.decode() for name in names if name)
    write(FOLDER / 'source-manifest.json', {n: sha(REPO / n) for n in paths if (REPO / n).is_file()})
    (FOLDER / 'source.patch').write_bytes(subprocess.check_output(
        ['git', 'diff', '--', 'crates', 'tools/perf-baseline'], cwd=REPO))
    print('Frozen baseline', flush=True)


def check_source():
    for name, digest in json.loads((FOLDER / 'source-manifest.json').read_text()).items():
        assert sha(REPO / name) == digest, name


def run(name, command, binary=None, allow_failure=False):
    receipt_path = FOLDER / (name + '.receipt.json')
    assert not receipt_path.exists(), receipt_path
    check_source()
    before = sha(binary) if binary else None
    start, tick = now(), time.monotonic()
    with (FOLDER / (name + '.stdout')).open('x') as out, (FOLDER / (name + '.stderr')).open('x') as err:
        result = subprocess.run(command, cwd=REPO, stdout=out, stderr=err)
    check_source()
    assert binary is None or sha(binary) == before
    receipt = dict(command=command, start_utc=start, end_utc=now(),
        seconds=time.monotonic()-tick, exit_code=result.returncode,
        binary_sha256=before, source_manifest_sha256=sha(FOLDER / 'source-manifest.json'),
        script_sha256=sha(Path(__file__)), plan_sha256=sha(HERE / 'plan.json'),
        environment={key: os.environ.get(key) for key in ['RUSTFLAGS',
            'CARGO_ENCODED_RUSTFLAGS', 'LD_PRELOAD', 'MALLOC_CONF', 'GLIBC_TUNABLES']})
    receipt['artifacts'] = {p.name: sha(p) for p in sorted(FOLDER.glob(name + '.*')) if p.is_file() and p != receipt_path}
    write(receipt_path, receipt)
    assert allow_failure or result.returncode == 0, (name, result.returncode)
    print(name, 'exit', result.returncode, flush=True)


def build(kind):
    executable = 'litchi-perf-baseline' + ('-alloc' if kind == 'alloc' else '')
    command = ['env', 'CARGO_BUILD_JOBS=2', 'CARGO_INCREMENTAL=0', 'cargo', 'build',
        '--release', '--locked', '--manifest-path', 'tools/perf-baseline/Cargo.toml',
        '--bin', executable, '--target-dir', str(TARGET)]
    if kind == 'alloc':
        command += ['--features', 'allocator-metrics']
    run('build-' + kind, command)
    source = TARGET / 'release' / executable
    binary = SCRATCH / kind
    shutil.copy2(source, binary)
    assert sha(source) == sha(binary)
    write(FOLDER / ('binary-' + kind + '.json'), dict(path=str(binary), sha256=sha(binary),
        bytes=binary.stat().st_size, build_receipt_sha256=sha(FOLDER / ('build-' + kind + '.receipt.json')),
        source_manifest_sha256=sha(FOLDER / 'source-manifest.json')))


def jobs(lane):
    plan = json.loads((HERE / 'plan.json').read_text())
    config = plan['allocation' if lane == 'alloc' else lane]
    output = []
    for repeat in range(1, config['repeats'] + 1):
        if lane == 'profile':
            selections = [('xls-owned', {'cases': ['xls_owned_source_open_one_cell']})]
            selections += [('cfb-' + s, {'cases': ['cfb_open'], 'shapes': [s],
                'payload': 'incompressible'}) for s in plan['groups']['cfb']['shapes']]
        elif lane == 'hardware':
            selections = [('xls-owned', {'cases': [config['case']]})]
        else:
            names = ['xls', 'cfb'] if repeat == 1 or lane == 'alloc' else ['cfb', 'xls']
            selections = [(n, plan['groups'][n]) for n in names]
        for group, selection in selections:
            output.append(dict(name=f'{lane}-r{repeat}-{group}', repeat=repeat,
                group=group, selection=selection, samples=config['samples'], warmup=config['warmup']))
    return output


def capture(lane):
    plan = json.loads((HERE / 'plan.json').read_text())
    kind = 'alloc' if lane == 'alloc' else 'normal'
    binary = SCRATCH / kind
    assert sha(binary) == json.loads((FOLDER / ('binary-' + kind + '.json')).read_text())['sha256']
    for job in jobs(lane):
        name, selection = job['name'], job['selection']
        command = ['taskset', '-c', str(plan['cpu'])]
        if lane == 'native':
            command += ['/usr/bin/time', '-f',
                '{"max_rss_kib":%M,"elapsed_seconds":%e,"user_seconds":%U,"system_seconds":%S}',
                '-o', str(FOLDER / (name + '.rss.json'))]
        elif lane == 'profile':
            owner = plan['profile']['cfb_owner' if job['group'].startswith('cfb-') else 'xls_owner']
            command += ['valgrind', '--tool=callgrind', '--collect-atstart=no',
                '--toggle-collect=' + owner, '--zero-before=' + owner, '--dump-after=' + owner,
                '--callgrind-out-file=' + str(FOLDER / (name + '.callgrind'))]
        elif lane == 'hardware':
            command += ['perf', 'stat', '-x', ',', '-o', str(FOLDER / (name + '.csv')),
                '-e', plan['hardware']['events'], '--']
        command += [str(binary), '--case', ','.join(selection['cases']),
            '--warmup', str(job['warmup']), '--samples', str(job['samples']),
            '--json', str(FOLDER / (name + '.json')),
            '--corpus-manifest', str(FOLDER / (name + '.catalog.json'))]
        if 'shapes' in selection:
            command += ['--shape', ','.join(selection['shapes']), '--payload', selection['payload']]
        run(name, command, binary, allow_failure=(lane == 'hardware'))


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('action', choices=['freeze', 'build-normal', 'build-alloc',
        'native', 'profile', 'alloc', 'hardware'])
    args = parser.parse_args()
    if args.action == 'freeze':
        freeze()
    elif args.action.startswith('build-'):
        build(args.action[6:])
    else:
        capture(args.action)
