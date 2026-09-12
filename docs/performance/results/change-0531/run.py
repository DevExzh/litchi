"""Serial, source-bound build and capture driver for the 0531 candidate."""
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
SCRATCH = Path('/tmp/litchi-goal-0531')
TARGET = Path('/home/zhuhe/litchi-goal-0531-target')


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def write(path, value):
    path.write_text(json.dumps(value, indent=2) + '\n')


def now():
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def plan_data():
    plan_path = HERE / 'plan.json'
    assert plan_path.is_file(), plan_path
    value = json.loads(plan_path.read_text())
    assert isinstance(value, dict), 'plan is not an object'
    return value


def candidate_roots():
    roots = plan_data().get('candidate_source_roots')
    assert isinstance(roots, list) and roots, \
        'plan candidate_source_roots must be a non-empty list'
    result = []
    for root in roots:
        assert isinstance(root, str) and root and not Path(root).is_absolute(), root
        path = Path(root.rstrip('/'))
        assert '..' not in path.parts and path.as_posix() == root.rstrip('/'), root
        assert root.startswith('crates/'), root
        assert (REPO / path).is_dir(), root
        result.append(root.rstrip('/') + '/')
    return tuple(result)


def candidate_files(before=None, after=None):
    """Return planned files, or derive and validate the frozen manifest diff."""

    files = plan_data().get('candidate_files', [])
    assert isinstance(files, list), 'plan candidate_files is not a list'
    planned = set()
    for name in files:
        assert isinstance(name, str) and name and not Path(name).is_absolute(), name
        path = Path(name)
        assert '..' not in path.parts and path.as_posix() == name, name
        assert any(name.startswith(root) for root in candidate_roots()), name
        planned.add(name)
    if before is None or after is None:
        return planned
    changed = {name for name in set(before) | set(after)
               if before.get(name) != after.get(name)}
    assert changed, 'candidate source diff is empty'
    roots = candidate_roots()
    assert all(any(name.startswith(root) for root in roots) for name in changed), \
        (sorted(changed), roots)
    if planned:
        assert changed == planned, (sorted(changed), sorted(planned))
    return changed


def source_paths():
    names = subprocess.check_output([
        'git', 'ls-files', '-z', 'crates', 'tools/perf-baseline', 'Cargo.toml',
        'Cargo.lock', '.cargo', 'rust-toolchain.toml'], cwd=REPO).split(b'\0')
    paths = {name.decode() for name in names if name}
    # A candidate may introduce a Rust module/test without staging it.  The
    # tracked-file listing alone would then let a build run without binding
    # the new source to the frozen receipt.  Include every non-ignored Rust
    # source under the measured crates and standalone harness, whether tracked
    # or not.  Git's ignored-file filter keeps generated target trees out of
    # the manifest.
    untracked = subprocess.check_output([
        'git', 'ls-files', '--others', '--exclude-standard', '-z', '--',
        'crates', 'tools/perf-baseline'], cwd=REPO).split(b'\0')
    paths.update(name.decode() for name in untracked
                 if name and name.endswith(b'.rs'))
    return paths


def manifest_paths():
    return {name for name in source_paths() if (REPO / name).is_file()}


def freeze(stage):
    assert stage in ('baseline', 'candidate')
    plan = plan_data()
    candidate_roots()
    assert plan.get('status') not in (None, 'draft-before-capture', 'draft'), \
        'plan status must be finalized before freeze'
    folder = HERE / stage
    folder.mkdir(exist_ok=False)
    paths = manifest_paths()
    write(folder/'source-manifest.json', {name: sha(REPO/name) for name in sorted(paths) if (REPO/name).is_file()})
    (folder/'source.patch').write_bytes(subprocess.check_output(['git', 'diff', '--', 'crates', 'tools/perf-baseline'], cwd=REPO))
    if stage == 'candidate':
        before = json.loads((HERE/'baseline/source-manifest.json').read_text())
        after = json.loads((folder/'source-manifest.json').read_text())
        changed = candidate_files(before, after)
        write(folder/'source-diff.json', {
            'baseline_manifest_sha256': sha(HERE/'baseline/source-manifest.json'),
            'candidate_manifest_sha256': sha(folder/'source-manifest.json'),
            'candidate_source_roots': list(candidate_roots()),
            'changed_files': {
                name: {'baseline_sha256': before.get(name),
                       'candidate_sha256': after.get(name)}
                for name in sorted(changed)
            },
        })
    print('Frozen', stage, flush=True)


def check_source(stage, retained_baseline=False):
    working_stage = stage
    if retained_baseline and stage == 'baseline' and (HERE/'candidate/source-manifest.json').exists():
        before = json.loads((HERE/'baseline/source-manifest.json').read_text())
        after = json.loads((HERE/'candidate/source-manifest.json').read_text())
        candidate_files(before, after)
        working_stage = 'candidate'
    manifest = json.loads((HERE/working_stage/'source-manifest.json').read_text())
    assert set(manifest) == manifest_paths(), (sorted(set(manifest) - manifest_paths()),
                                                sorted(manifest_paths() - set(manifest)))
    for name, digest in manifest.items():
        assert sha(REPO/name) == digest, name
    return working_stage


def run(stage, name, command, binary=None, retained_baseline=False):
    folder = HERE/stage
    receipt_path = folder/(name+'.receipt.json')
    assert not receipt_path.exists(), receipt_path
    working_stage = check_source(stage, retained_baseline)
    before = sha(binary) if binary else None
    start, tick = now(), time.monotonic()
    test_tmp = TARGET / 'test-tmp'
    test_tmp.mkdir(parents=True, exist_ok=True)
    environment = dict(os.environ)
    environment['TMPDIR'] = str(test_tmp)
    with (folder/(name+'.stdout')).open('w') as out, (folder/(name+'.stderr')).open('w') as err:
        result = subprocess.run(command, cwd=REPO, stdout=out, stderr=err, env=environment)
    assert check_source(stage, retained_baseline) == working_stage
    assert binary is None or sha(binary) == before
    receipt = dict(command=command, start_utc=start, end_utc=now(), seconds=time.monotonic()-tick,
                   exit_code=result.returncode, binary_sha256=before,
                   source_manifest_sha256=sha(folder/'source-manifest.json'),
                   working_source_manifest_sha256=sha(HERE/working_stage/'source-manifest.json'),
                   script_sha256=sha(Path(__file__)), plan_sha256=sha(HERE/'plan.json'),
                   environment={key: environment.get(key) for key in ['RUSTFLAGS', 'CARGO_ENCODED_RUSTFLAGS', 'LD_PRELOAD', 'MALLOC_CONF', 'GLIBC_TUNABLES', 'TMPDIR']})
    receipt['artifacts'] = {p.name: sha(p) for p in sorted(folder.glob(name+'.*')) if p.is_file() and p != receipt_path}
    write(receipt_path, receipt)
    assert result.returncode == 0, (name, result.returncode)
    print(stage, name, 'passed', flush=True)


def build(stage, kind):
    SCRATCH.mkdir(parents=True, exist_ok=True)
    executable = 'litchi-perf-baseline' + ('-alloc' if kind == 'alloc' else '')
    command = ['env', 'CARGO_BUILD_JOBS=2', 'CARGO_INCREMENTAL=0', 'cargo', 'build', '--release', '--locked', '--manifest-path', 'tools/perf-baseline/Cargo.toml',
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


def capture(stage, lane, selected_repeat=None):
    plan = plan_data()
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
        config = plan['allocation' if lane == 'alloc' else
                       ('profile' if lane == 'profile' else 'hardware')]
        for repeat in range(1, config['repeats']+1):
            for shape in config['shapes']:
                jobs.append((f'{lane}-r{repeat}-{shape}', plan['primary']['case'], shape, config['warmup'], config['samples']))
    for name, case, shape, warmup, samples in jobs:
        if selected_repeat is not None and f'-r{selected_repeat}-' not in name:
            continue
        command = ['taskset', '-c', str(plan['cpu'])]
        if lane == 'native':
            command += ['/usr/bin/time', '-f', '{"max_rss_kib":%M,"elapsed_seconds":%e,"user_seconds":%U,"system_seconds":%S}',
                        '-o', str(HERE/stage/(name+'.rss.json'))]
        elif lane == 'profile':
            owner = plan['profile']['owner']
            command += ['valgrind', '--tool=callgrind', '--collect-atstart=no', '--toggle-collect='+owner,
                        '--zero-before='+owner, '--dump-after='+owner,
                        '--callgrind-out-file='+str(HERE/stage/(name+'.callgrind'))]
        elif lane == 'hardware':
            command += ['perf', 'stat', '-x', ',', '-o', str(HERE/stage/(name+'.csv')),
                        '-e', plan['hardware']['events'], '--']
        command += [str(binary), '--warmup', str(warmup), '--samples', str(samples), '--case', case,
                    '--xlsx-cell-crud-shape', shape, '--json', str(HERE/stage/(name+'.json'))]
        run(stage, name, command, binary, retained_baseline=(stage == 'baseline'))


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('stage', choices=['baseline', 'candidate'])
    parser.add_argument('action', choices=['freeze', 'build-normal', 'build-alloc',
                                           'native', 'native-r1', 'native-r2',
                                           'profile', 'alloc', 'hardware'])
    args = parser.parse_args()
    if args.action == 'freeze':
        freeze(args.stage)
    elif args.action.startswith('build-'):
        build(args.stage, args.action[6:])
    elif args.action.startswith('native-r'):
        capture(args.stage, 'native', int(args.action[-1]))
    else:
        capture(args.stage, args.action)
