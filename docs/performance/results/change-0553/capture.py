"""Frozen serial baseline/candidate capture matrix; immutable binary copies."""

import argparse
import json
import shutil

import run as R


def plan():
    return json.loads((R.HERE / 'plan.json').read_text())


def main_lane(lane, repeat):
    p = plan()
    kind = 'alloc' if lane == 'alloc' else 'normal'
    binary = R.SCRATCH / kind
    identity = json.loads((R.FOLDER / ('binary-' + kind + '.json')).read_text())
    assert R.sha(binary) == identity['sha256']
    for shape in p['shapes'][::1 if repeat == 1 else -1]:
        for index, case in enumerate(p['cases']):
            name = f'{lane}-r{repeat}-{shape}-c{index}'
            samples = 1 if lane == 'preflight' else p[lane]['samples']
            warmup = 0 if lane == 'preflight' else p[lane]['warmup']
            command = ['taskset', '-c', str(p['cpu']), '/usr/bin/time', '-v',
                       str(binary), '--case', case, '--xlsx-cell-crud-shape', shape,
                       '--samples', str(samples), '--warmup', str(warmup),
                       '--json', str(R.FOLDER / (name + '.json')),
                       '--corpus-manifest', str(R.FOLDER / (name + '.catalog.json'))]
            R.run(name, command, binary)


def build_guard(kind):
    command = ['env', 'TMPDIR=' + str(R.TARGET / 'tmp'), 'CARGO_BUILD_JOBS=2',
               'CARGO_INCREMENTAL=0', 'cargo', 'build', '--release', '--locked',
               '--manifest-path', 'tools/perf-baseline/Cargo.toml', '--bin',
               'xlsx_planning_guard', '--target-dir', str(R.TARGET)]
    if kind == 'alloc':
        command += ['--features', 'allocator-metrics']
    R.run('build-guard-' + kind, command)
    binary = R.SCRATCH / ('guard-' + kind)
    shutil.copy2(R.TARGET / 'release/xlsx_planning_guard', binary)
    R.write(R.FOLDER / ('binary-guard-' + kind + '.json'), {
        'path': str(binary), 'sha256': R.sha(binary), 'bytes': binary.stat().st_size,
        'build_receipt_sha256': R.sha(R.FOLDER / ('build-guard-' + kind + '.receipt.json')),
        'source_manifest_sha256': R.sha(R.FOLDER / 'source-manifest.json'),
    })


def guard_lane(lane, repeat):
    p = plan()
    config = p['guard']
    kind = 'alloc' if lane == 'alloc' else 'normal'
    binary = R.SCRATCH / ('guard-' + kind)
    identity = json.loads((R.FOLDER / ('binary-guard-' + kind + '.json')).read_text())
    assert R.sha(binary) == identity['sha256']
    for shape in config['shapes'][::1 if repeat == 1 else -1]:
        for case in config['cases']:
            name = f'guard-{lane}-r{repeat}-{shape}-{case}'
            command = ['taskset', '-c', str(p['cpu']), str(binary), '--shape', shape,
                       '--case', case, '--samples', str(config[lane + '_samples']),
                       '--warmup', str(config[lane + '_warmup']),
                       '--json', str(R.FOLDER / (name + '.json'))]
            R.run(name, command, binary)


def build_cap():
    command = ['env', 'TMPDIR=' + str(R.TARGET / 'tmp'), 'CARGO_BUILD_JOBS=2',
               'CARGO_INCREMENTAL=0', 'cargo', 'build', '--release', '--locked',
               '-p', 'litchi-xlsx', '--example', 'perf_cap_boundary',
               '--target-dir', str(R.TARGET)]
    R.run('build-cap', command)
    binary = R.SCRATCH / 'cap'
    shutil.copy2(R.TARGET / 'release/examples/perf_cap_boundary', binary)
    R.write(R.FOLDER / 'binary-cap.json', {
        'path': str(binary), 'sha256': R.sha(binary), 'bytes': binary.stat().st_size,
        'build_receipt_sha256': R.sha(R.FOLDER / 'build-cap.receipt.json'),
        'source_manifest_sha256': R.sha(R.FOLDER / 'source-manifest.json'),
    })


def cap_lane(repeat):
    p = plan()
    binary = R.SCRATCH / 'cap'
    assert R.sha(binary) == json.loads((R.FOLDER / 'binary-cap.json').read_text())['sha256']
    for size in p['cap']['sizes'][::1 if repeat == 1 else -1]:
        name = f'cap-r{repeat}-{size}'
        R.run(name, ['taskset', '-c', str(p['cpu']), str(binary), '--size', str(size),
                    '--samples', str(p['cap']['samples']),
                    '--warmup', str(p['cap']['warmup']),
                    '--json', str(R.FOLDER / (name + '.json')),
                    '--fixture-out', str(R.FOLDER / (name + '.zip'))], binary)


def profiles(repeat):
    p = plan()
    binary = R.SCRATCH / 'normal'
    assert R.sha(binary) == json.loads((R.FOLDER / 'binary-normal.json').read_text())['sha256']
    for shape in p['shapes'][::1 if repeat == 1 else -1]:
        name = f'profile-r{repeat}-{shape}-c0'
        owner = p['profile']['owner']
        R.run(name, ['taskset', '-c', str(p['cpu']), 'valgrind', '--tool=callgrind',
                    '--vgdb=no', '--vgdb-prefix=' + str(R.TARGET / 'tmp/vgdb'),
                    '--collect-atstart=no', '--toggle-collect=' + owner,
                    '--zero-before=' + owner, '--dump-after=' + owner,
                    '--callgrind-out-file=' + str(R.FOLDER / (name + '.callgrind')),
                    str(binary), '--case', p['profile']['case'],
                    '--xlsx-cell-crud-shape', shape, '--samples', '1', '--warmup', '0',
                    '--json', str(R.FOLDER / (name + '.json')),
                    '--corpus-manifest', str(R.FOLDER / (name + '.catalog.json'))], binary)


def campaign(stage):
    R.configure(stage)
    if stage == 'baseline':
        (R.TARGET / 'tmp').mkdir(parents=True, exist_ok=False)
    R.freeze()
    repeats = (1,) if stage == 'baseline' else (1, 2)
    R.build('normal')
    main_lane('preflight', 1)
    for repeat in repeats:
        main_lane('native', repeat)
    build_guard('normal')
    for repeat in repeats:
        guard_lane('native', repeat)
    build_cap()
    for repeat in repeats:
        cap_lane(repeat)
    R.build('alloc')
    for repeat in repeats:
        main_lane('alloc', repeat)
    build_guard('alloc')
    for repeat in repeats:
        guard_lane('alloc', repeat)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('action', choices=('baseline', 'candidate', 'baseline-r2', 'profiles'))
    args = parser.parse_args()
    if args.action in ('baseline', 'candidate'):
        campaign(args.action)
    elif args.action == 'baseline-r2':
        R.configure('baseline', 'candidate')
        main_lane('native', 2)
        guard_lane('native', 2)
        cap_lane(2)
        main_lane('alloc', 2)
        guard_lane('alloc', 2)
    else:
        for stage in ('baseline', 'candidate'):
            R.configure(stage, 'candidate')
            for repeat in (1, 2):
                profiles(repeat)
