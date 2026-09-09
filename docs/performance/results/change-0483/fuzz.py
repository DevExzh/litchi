#!/usr/bin/env python3
"""Prepare, build, and retain a bounded ASan/libFuzzer DOCX tail-append smoke run."""
import json
from pathlib import Path
import shutil
import subprocess
import sys

from common import ROOT, REPO, TEMP, ENV, meta, now, read, write
from gate import snapshot

FUZZ = ROOT / 'fuzz'
ATTEMPT = sys.argv[2] if len(sys.argv) > 2 else 'accepted'
assert ATTEMPT and '/' not in ATTEMPT and ATTEMPT not in ('.', '..')
DATA = FUZZ if ATTEMPT == 'initial' else FUZZ / ATTEMPT
WORK = TEMP / 'fuzz-docx' if ATTEMPT == 'initial' else TEMP / f'fuzz-docx-{ATTEMPT}'
TARGET = 'x86_64-unknown-linux-gnu'
FLAGS = ('-C passes=sancov-module -C llvm-args=-sanitizer-coverage-level=4 '
         '-C llvm-args=-sanitizer-coverage-inline-8bit-counters '
         '-C llvm-args=-sanitizer-coverage-pc-table '
         '-C llvm-args=-sanitizer-coverage-trace-compares -Z sanitizer=address --cfg fuzzing')


def inventory(directory):
    return {p.relative_to(directory).as_posix(): meta(p)
            for p in sorted(directory.rglob('*')) if p.is_file()}


def prepare():
    WORK.mkdir(parents=True, exist_ok=False)
    DATA.mkdir(parents=True, exist_ok=True)
    original = REPO / 'crates/litchi-docx/fuzz/Cargo.toml'
    target = REPO / 'crates/litchi-docx/fuzz/fuzz_targets/tail_append.rs'
    manifest = original.read_text().replace(
        'path = ".."', 'path = ' + json.dumps(str(REPO / 'crates/litchi-docx'))
    ).replace('path = "fuzz_targets/tail_append.rs"', 'path = ' + json.dumps(str(target))).replace(
        'path = "fuzz_targets/parse_docx.rs"', 'path = ' + json.dumps(str(REPO / 'crates/litchi-docx/fuzz/fuzz_targets/parse_docx.rs')))
    with (WORK / 'Cargo.toml').open('x') as stream:
        stream.write(manifest)
    inputs = DATA / 'build-inputs'
    inputs.mkdir(exist_ok=False)
    with (inputs / 'Cargo.toml').open('x') as stream:
        stream.write(manifest)
    subprocess.run(['cargo', 'generate-lockfile', '--offline', '--manifest-path',
                    str(WORK / 'Cargo.toml')], cwd=REPO, env=ENV, check=True)
    shutil.copyfile(WORK / 'Cargo.lock', inputs / 'Cargo.lock.txt')
    corpus = WORK / 'corpus'
    corpus.mkdir()
    seeds = inventory(FUZZ / 'seeds')
    for name in seeds:
        shutil.copyfile(FUZZ / 'seeds' / name, corpus / name.replace('/', '-'))
    write(DATA / 'prepared.json', {
        'prepared_utc': now(), 'target_source': meta(target),
        'manifest': meta(WORK / 'Cargo.toml'), 'lock': meta(WORK / 'Cargo.lock'),
        'seeds': seeds, 'corpus_before': inventory(corpus),
    })


def verify_inputs():
    prepared = read(DATA / 'prepared.json')
    assert meta(WORK / 'Cargo.toml') == prepared['manifest']
    assert meta(WORK / 'Cargo.lock') == prepared['lock']
    assert meta(REPO / 'crates/litchi-docx/fuzz/fuzz_targets/tail_append.rs') == prepared['target_source']
    assert inventory(FUZZ / 'seeds') == prepared['seeds']
    return prepared


def build():
    prepared = verify_inputs()
    source = snapshot()
    argv = ['env', f'CARGO_TARGET_DIR={REPO / "target/fuzz-asan"}', 'RUSTC_BOOTSTRAP=1',
            f'RUSTFLAGS={FLAGS}', 'cargo', 'build', '--release', '--locked',
            '--manifest-path', str(WORK / 'Cargo.toml'), '--target', TARGET, '--bin', 'tail_append']
    print(json.dumps(argv), flush=True)
    started = now()
    subprocess.run(argv, cwd=REPO, env=ENV, check=True)
    assert snapshot() == source
    origin = REPO / 'target/fuzz-asan' / TARGET / 'release/tail_append'
    destination = WORK / 'tail_append'
    with origin.open('rb') as src, destination.open('xb') as dst:
        shutil.copyfileobj(src, dst)
    destination.chmod(0o755)
    assert meta(origin) == meta(destination)
    write(DATA / 'build.json', {
        'argv': argv, 'cwd': str(REPO), 'started_utc': started, 'finished_utc': now(),
        'source_snapshot': source, 'inputs': prepared,
        'binary': {'path': str(destination), **meta(destination)},
    })


def smoke():
    verify_inputs()
    build_record = read(DATA / 'build.json')
    binary = Path(build_record['binary']['path'])
    assert meta(binary) == {k: build_record['binary'][k] for k in ('bytes', 'sha256')}
    corpus = WORK / 'corpus'
    artifacts = WORK / 'artifacts'
    artifacts.mkdir(exist_ok=False)
    before = inventory(corpus)
    argv = [str(binary), str(corpus), '-runs=10000', '-seed=483', '-max_len=65536',
            '-timeout=10', f'-artifact_prefix={artifacts}/']
    started = now()
    print(json.dumps(argv), flush=True)
    result = subprocess.run(argv, cwd=REPO, env=ENV)
    # Preserve mutated corpus and crash artifacts for every terminal outcome.
    retained = DATA / 'post-run'
    retained.mkdir(exist_ok=False)
    shutil.copytree(corpus, retained / 'corpus')
    shutil.copytree(artifacts, retained / 'artifacts')
    after_binary = meta(binary)
    write(DATA / 'smoke.json', {
        'argv': argv, 'cwd': str(REPO), 'started_utc': started, 'finished_utc': now(),
        'exit_code': result.returncode, 'binary_before': build_record['binary'],
        'binary_after': {'path': str(binary), **after_binary},
        'corpus_before': before, 'retained': inventory(retained),
    })
    assert after_binary == {k: build_record['binary'][k] for k in ('bytes', 'sha256')}
    raise SystemExit(result.returncode)


if __name__ == '__main__':
    {'prepare': prepare, 'build': build, 'smoke': smoke}[sys.argv[1]]()
