#!/usr/bin/env python3
"""Build and profile the committed single-scan XLSX parser in a clean tree."""
import datetime
import gzip
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
TEMP = Path('/tmp/litchi-goal-0468')
TREE = TEMP / 'profile-tree'
BINARY = TEMP / 'profile'
ENV = dict(os.environ, RUSTUP_TOOLCHAIN='1.98.1', CARGO_BUILD_JOBS='4',
           CARGO_INCREMENTAL='0', CARGO_PROFILE_RELEASE_DEBUG='1',
           RUSTFLAGS='-C force-frame-pointers=yes -C force-unwind-tables=yes',
           DEBUGINFOD_URLS='', LC_ALL='C', PYTHONDONTWRITEBYTECODE='1')


def sha(path):
    with path.open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()


def write(path, value):
    with path.open('x') as stream:
        json.dump(value, stream, indent=2, sort_keys=True)
        stream.write('\n')


def check_source():
    binding = json.loads((ROOT / 'source-binding.json').read_text())
    revision = subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=TREE, text=True).strip()
    assert revision == binding['revision']
    assert not subprocess.check_output(['git', 'status', '--porcelain'], cwd=TREE)
    assert sha(ROOT / 'sources.json') == binding['source_manifest_sha256']
    for name, digest in json.loads((ROOT / 'sources.json').read_text()).items():
        assert sha(TREE / name) == digest, name
    for name, digest in binding['included_fixtures'].items():
        assert sha(TREE / name) == digest, name
    return binding


def run(argv, prefix):
    receipt = dict(schema='litchi-0468-command-v1', argv=argv, cwd=str(TREE),
                   driver_sha256=sha(Path(__file__)),
                   environment={k: ENV[k] for k in ('RUSTUP_TOOLCHAIN', 'RUSTFLAGS',
                       'CARGO_PROFILE_RELEASE_DEBUG', 'CARGO_BUILD_JOBS',
                       'CARGO_INCREMENTAL', 'DEBUGINFOD_URLS', 'LC_ALL')},
                   started_utc=datetime.datetime.now(datetime.timezone.utc).isoformat())
    write(prefix.with_suffix('.started.json'), receipt)
    with prefix.with_suffix('.stdout').open('xb') as out, prefix.with_suffix('.stderr').open('xb') as err:
        p = subprocess.run(argv, cwd=TREE, env=ENV, stdout=out, stderr=err)
    receipt.update(exit_code=p.returncode,
                   finished_utc=datetime.datetime.now(datetime.timezone.utc).isoformat())
    receipt['outputs'] = {suffix: dict(sha256=sha(prefix.with_suffix(suffix)),
                         bytes=prefix.with_suffix(suffix).stat().st_size)
                         for suffix in ('.stdout', '.stderr')}
    write(prefix.with_suffix('.json'), receipt)
    print(prefix.name, p.returncode, flush=True)
    if p.returncode:
        raise SystemExit(p.returncode)


def build():
    source = check_source()
    run(['cargo', 'build', '--release', '--locked', '--manifest-path',
         'tools/perf-baseline/Cargo.toml', '--target-dir',
         str(REPO / 'tools/perf-baseline/target'), '--bin', 'litchi-perf-baseline'], ROOT / 'build')
    check_source()
    assert not BINARY.exists()
    shutil.copy2(REPO / 'tools/perf-baseline/target/release/litchi-perf-baseline', BINARY)
    write(ROOT / 'binding.json', dict(schema='litchi-0468-build-binding-v1',
          revision=source['revision'], clean_build=True, binary_sha256=sha(BINARY),
          bytes=BINARY.stat().st_size, binary_path=str(BINARY),
          build_receipt_sha256=sha(ROOT / 'build.json'),
          source_binding_sha256=sha(ROOT / 'source-binding.json')))


def capture(lane):
    source = check_source()
    binding = json.loads((ROOT / 'binding.json').read_text())
    assert sha(BINARY) == binding['binary_sha256']
    protocol = json.loads((ROOT / 'protocol.json').read_text())
    assert protocol['capture_driver_sha256'] == sha(Path(__file__))
    config = protocol['lanes'][lane]
    output = ROOT / lane
    output.mkdir()
    command = ['taskset', '-c', '2', '/usr/bin/time', '-v', '-o', str(output / 'resource.log')]
    if lane == 'counters':
        command += ['perf', 'stat', '-x', ';', '-o', str(output / 'counters.csv'),
                    '-e', 'cycles:u,instructions:u,branches:u,branch-misses:u,cache-misses:u,page-faults', '--']
    if lane == 'samples-fp':
        command += ['perf', 'record', '-F', '499', '-e', 'cycles:u', '--call-graph', 'fp',
                    '-o', str(output / 'perf.data'), '--']
    command += [str(BINARY), '--case', 'xlsx_one_percent_commit_save', '--xlsx-shape', 'dense-wide',
                '--workers', '1', '--warmup', str(config['warmups']), '--samples', str(config['samples']),
                '--json', str(output / 'report.json'), '--corpus-manifest', str(output / 'corpus-catalog.json')]
    run(command, output / 'capture')
    check_source()
    assert sha(BINARY) == binding['binary_sha256']
    report = json.loads((output / 'report.json').read_text())
    assert report['environment']['git_revision'] == source['revision']
    assert report['environment']['git_worktree_dirty'] is False
    assert len(report['results']) == 1
    row = report['results'][0]
    assert row['case'] == 'xlsx_one_percent_commit_save'
    assert row['corpus']['archive_sha256'] == protocol['corpus_sha256']
    assert len(row['elapsed_ns']['samples']) == config['samples']
    write(output / 'receipt.json', dict(schema='litchi-0468-capture-v1', lane=lane,
          revision=source['revision'], clean_source_before_and_after=True,
          source_and_binary_unchanged=True, binary_sha256=binding['binary_sha256'],
          binding_sha256=sha(ROOT / 'binding.json'), protocol_sha256=sha(ROOT / 'protocol.json'),
          artifacts={p.name: dict(sha256=sha(p), bytes=p.stat().st_size)
                     for p in sorted(output.iterdir()) if p.is_file()}))


def export():
    check_source()
    assert sha(BINARY) == json.loads((ROOT / 'binding.json').read_text())['binary_sha256']
    output = ROOT / 'samples-fp'
    run(['perf', 'script', '--no-inline', '-i', str(output / 'perf.data'),
         '-F', 'comm,pid,tid,time,event,period,ip,sym,dso'], output / 'perf-script')
    run(['perf', 'report', '--stdio', '--no-children', '--no-inline', '-g', 'none',
         '-i', str(output / 'perf.data'), '--sort', 'symbol', '--percent-limit', '0.5'], output / 'top-symbols')
    compressed = []
    for name in ('perf.data', 'perf-script.stdout'):
        path = output / name
        target = output / (name + '.gz')
        assert not target.exists()
        with path.open('rb') as src, target.open('xb') as dst:
            with gzip.GzipFile(filename='', fileobj=dst, mode='wb', mtime=0) as stream:
                shutil.copyfileobj(src, stream)
        with gzip.open(target, 'rb') as stream:
            assert hashlib.file_digest(stream, 'sha256').hexdigest() == sha(path)
        compressed.append(dict(path=str(path.relative_to(ROOT)), sha256=sha(path), bytes=path.stat().st_size,
                          compressed_path=str(target.relative_to(ROOT)), compressed_sha256=sha(target),
                          compressed_bytes=target.stat().st_size))
        path.unlink()
    write(ROOT / 'compression.json', dict(schema='litchi-0468-compression-v1', artifacts=compressed))


if __name__ == '__main__':
    action = sys.argv[1]
    if action == 'build':
        build()
    elif action == 'export':
        export()
    else:
        capture(action)
