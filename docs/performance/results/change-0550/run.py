"""Serial source-bound XLSX commit attribution driver; no runtime changes."""
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
SCRATCH_ROOT = Path('/home/zhuhe/litchi-goal-0550-target/retained')
SCRATCH = SCRATCH_ROOT / 'baseline'
STAGE = 'baseline'
EXECUTION_STAGE = 'baseline'
TARGET = Path('/home/zhuhe/litchi-goal-0550-target')
FOLDER = HERE / 'baseline'


def configure(stage, execution_stage=None):
    global STAGE, EXECUTION_STAGE, FOLDER, SCRATCH
    assert stage in ('baseline', 'candidate', 'final')
    STAGE = stage
    EXECUTION_STAGE = execution_stage or stage
    assert EXECUTION_STAGE in ('baseline', 'candidate', 'final')
    FOLDER = HERE / stage
    SCRATCH = SCRATCH_ROOT / stage


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def write(path, value):
    with path.open('x') as stream:
        stream.write(json.dumps(value, indent=2) + '\n')


def now():
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def freeze():
    FOLDER.mkdir(exist_ok=False)
    SCRATCH.mkdir(parents=True, exist_ok=True)
    names = subprocess.check_output(['git', 'ls-files', '-z', 'crates',
        'tools/perf-baseline', 'Cargo.toml', 'Cargo.lock', '.cargo',
        'rust-toolchain.toml'], cwd=REPO).split(b'\0')
    untracked = subprocess.check_output(['git', 'ls-files', '--others', '--exclude-standard', '-z', '--', 'crates', 'tools/perf-baseline'], cwd=REPO).split(b'\0')
    paths = sorted({name.decode() for name in names if name} | {name.decode() for name in untracked if name.endswith(b'.rs')})
    write(FOLDER / 'source-manifest.json', {n: sha(REPO / n) for n in paths if (REPO / n).is_file()})
    (FOLDER / 'source.patch').write_bytes(subprocess.check_output(
        ['git', 'diff', '--', 'crates', 'tools/perf-baseline'], cwd=REPO))
    print('Frozen', STAGE, flush=True)


def check_source():
    for name, digest in json.loads((HERE / EXECUTION_STAGE / 'source-manifest.json').read_text()).items():
        assert sha(REPO / name) == digest, name


def run(name, command, binary=None, allow_failure=False):
    receipt_path = FOLDER / (name + '.receipt.json')
    assert not receipt_path.exists(), receipt_path
    check_source()
    before = sha(binary) if binary else None
    observations = []
    for proc in Path('/proc').iterdir():
        if proc.name.isdigit():
            try:
                comm = (proc / 'comm').read_text().strip()
                if comm in ('cargo', 'rustc'):
                    observations.append(dict(pid=int(proc.name), comm=comm, cwd=os.readlink(proc / 'cwd')))
            except OSError:
                pass
    write(FOLDER / (name + '.host.json'), dict(observed_utc=now(), compiler_processes=observations, scope='Accessible compiler processes; no host quiescence guarantee'))
    start, tick = now(), time.monotonic()
    with (FOLDER / (name + '.stdout')).open('x') as out, (FOLDER / (name + '.stderr')).open('x') as err:
        result = subprocess.run(command, cwd=REPO, stdout=out, stderr=err)
    check_source()
    assert binary is None or sha(binary) == before
    receipt = dict(command=command, start_utc=start, end_utc=now(),
        seconds=time.monotonic()-tick, exit_code=result.returncode,
        execution_stage=EXECUTION_STAGE, execution_manifest_sha256=sha(HERE / EXECUTION_STAGE / 'source-manifest.json'),
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
    command = ['env', 'TMPDIR=' + str(TARGET / 'tmp'), 'CARGO_BUILD_JOBS=2', 'CARGO_INCREMENTAL=0', 'cargo', 'build',
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

