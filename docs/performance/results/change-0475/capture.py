#!/usr/bin/env python3
"""Authenticate and profile the unchanged public PPTX streaming benchmark."""
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
TEMP = Path('/tmp/litchi-goal-0475')
# Recreate the original source path to retain exact DWARF/source and runtime Git identity.
TREE_ROOT = Path('/tmp/litchi-goal-0474')
TREE = TREE_ROOT / 'tree'
BINARY = TEMP / 'profile'
ENV = dict(os.environ, RUSTUP_TOOLCHAIN='1.98.1', CARGO_PROFILE_RELEASE_DEBUG='1',
           RUSTFLAGS='-C force-frame-pointers=yes -C force-unwind-tables=yes',
           DEBUGINFOD_URLS='', LC_ALL='C', PYTHONDONTWRITEBYTECODE='1')
ENV_KEYS = ('RUSTUP_TOOLCHAIN', 'CARGO_PROFILE_RELEASE_DEBUG', 'RUSTFLAGS', 'DEBUGINFOD_URLS', 'LC_ALL')

def sha(path):
    with path.open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()

def write(path, value):
    with path.open('x') as stream:
        json.dump(value, stream, indent=2, sort_keys=True); stream.write('\n')

def now():
    return datetime.datetime.now(datetime.timezone.utc).isoformat()

def check():
    binding = json.loads((ROOT / 'binding.json').read_text())
    assert sha(BINARY) == binding['binary']['sha256']
    assert subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=TREE, text=True).strip() == binding['revision']
    assert not subprocess.check_output(['git', 'status', '--porcelain'], cwd=TREE).strip()
    assert sha(ROOT / 'reuse/sources/source.json') == binding['source_manifest_sha256']
    for path, digest in json.loads((ROOT / 'reuse/sources/source.json').read_text()).items():
        assert sha(TREE / path) == digest, path
    for path, digest in binding['fixtures'].items():
        assert sha(TREE / path) == digest, path
    return binding

def command(argv, prefix, **extra):
    started = dict(argv=argv, cwd=str(TREE), environment={k: ENV[k] for k in ENV_KEYS},
                   started_utc=now(), driver_sha256=sha(Path(__file__)), **extra)
    write(prefix.with_suffix('.started.json'), started)
    with prefix.with_suffix('.stdout').open('xb') as out, prefix.with_suffix('.stderr').open('xb') as err:
        result = subprocess.run(argv, cwd=TREE, env=ENV, stdout=out, stderr=err)
    receipt = dict(started, exit_code=result.returncode, finished_utc=now(),
                   artifacts={prefix.with_suffix(s).name: dict(sha256=sha(prefix.with_suffix(s)), bytes=prefix.with_suffix(s).stat().st_size)
                              for s in ('.stdout', '.stderr')})
    write(prefix.with_suffix('.json'), receipt)
    print(prefix.parent.name, prefix.name, result.returncode, flush=True)
    return result.returncode

def prepare():
    protocol = json.loads((ROOT / 'protocol.json').read_text())
    assert protocol['capture_driver_sha256'] == sha(Path(__file__))
    assert not TREE_ROOT.exists() and not TEMP.exists()
    prior = json.loads((ROOT / 'reuse/build.json').read_text())
    cached = REPO / 'tools/perf-baseline/target/release/litchi-perf-baseline'
    assert sha(cached) == prior['binaries']['normal']['sha256']
    assert cached.stat().st_size == prior['binaries']['normal']['bytes']
    TEMP.mkdir(); TREE_ROOT.mkdir(); (TEMP / 'cpu.lock').touch()
    started = now()
    subprocess.run(['git', 'worktree', 'add', '--detach', '--no-checkout', str(TREE), prior['revision']], cwd=REPO, check=True)
    tracked = subprocess.check_output(['git', 'ls-tree', '-r', '--name-only', prior['revision']], cwd=REPO, text=True).splitlines()
    source = json.loads((ROOT / 'reuse/sources/source.json').read_text())
    selected = [p for p in tracked if not p.startswith('docs/') or p in source]
    subprocess.run(['git', 'sparse-checkout', 'set', '--no-cone', '--stdin'], cwd=TREE,
                   input=''.join('/' + p + '\n' for p in selected), text=True, check=True)
    subprocess.run(['git', 'reset', '--hard', prior['revision']], cwd=TREE, check=True)
    for path, digest in prior['fixtures'].items():
        assert sha(REPO / path) == digest
        target = TREE / path; target.parent.mkdir(parents=True, exist_ok=True); shutil.copy2(REPO / path, target)
    shutil.copy2(cached, BINARY)
    write(ROOT / 'binding.json', dict(schema='litchi-0475-reused-build-v1', revision=prior['revision'],
          tree=str(TREE), source_manifest_sha256=sha(ROOT / 'reuse/sources/source.json'),
          source_files=len(source), fixtures=prior['fixtures'],
          binary=dict(path=str(BINARY), bytes=BINARY.stat().st_size, sha256=sha(BINARY)),
          reused_build_sha256=sha(ROOT / 'reuse/build.json'), cached_binary_path=str(cached),
          protocol_sha256=sha(ROOT / 'protocol.json'), capture_driver_sha256=sha(Path(__file__)),
          prepared_started_utc=started, prepared_finished_utc=now(), fresh_build=False))
    check()
    print('authenticated reuse prepared', flush=True)

def capture(lane):
    binding = check()
    protocol = json.loads((ROOT / 'protocol.json').read_text())
    assert protocol['capture_driver_sha256'] == sha(Path(__file__))
    item = next(x for x in protocol['order'] if x['lane'] == lane)
    output = ROOT / lane; output.mkdir(exist_ok=False)
    argv = ['taskset', '-c', '2', '/usr/bin/time', '-v', '-o', str(output / 'resource.log')]
    if item['kind'] == 'counters':
        argv += ['perf', 'stat', '-x', ';', '-o', str(output / 'counters.csv'), '-e', protocol['counter_events'], '--']
    elif item['kind'] == 'cpu':
        argv += ['perf', 'record', '-F', '499', '-e', 'cycles:u', '--call-graph', 'fp', '-o', str(output / 'perf.data'), '--']
    elif item['kind'] == 'heap':
        argv += ['heaptrack', '--record-only', '-o', str(output / 'heaptrack')]
    argv += [str(BINARY), '--case', 'pptx_streaming_create', '--semantic-shape', 'large', '--workers', '1',
             '--warmup', str(item['warmups']), '--samples', str(item['samples']),
             '--json', str(output / 'report.json'), '--corpus-manifest', str(output / 'corpus-catalog.json')]
    code = command(argv, output / 'capture', lane=lane, binding_sha256=sha(ROOT / 'binding.json'),
                   protocol_sha256=sha(ROOT / 'protocol.json'), binary_sha256=sha(BINARY), clean_before=True)
    check()
    artifacts = {p.name: dict(sha256=sha(p), bytes=p.stat().st_size) for p in sorted(output.iterdir()) if p.is_file()}
    write(output / 'receipt.json', dict(schema='litchi-0475-capture-v1', lane=lane, kind=item['kind'],
          exit_code=code, clean_after=True, binary_unchanged=True, source_unchanged=True,
          binding_sha256=sha(ROOT / 'binding.json'), protocol_sha256=sha(ROOT / 'protocol.json'), artifacts=artifacts))
    if code:
        raise SystemExit(code)
    report = json.loads((output / 'report.json').read_text())
    assert report['environment']['git_revision'] == binding['revision']
    assert report['environment']['git_worktree_dirty'] is False
    assert len(report['results']) == 1
    row = report['results'][0]
    assert row['case'] == 'pptx_streaming_create'
    assert row['corpus']['archive_sha256'] == protocol['corpus_sha256'] == row['output_sha256']
    assert len(row['elapsed_ns']['samples']) == item['samples']

def compress(path):
    target = path.with_name(path.name + '.gz')
    with path.open('rb') as src, target.open('xb') as dst:
        with gzip.GzipFile(filename='', fileobj=dst, mode='wb', mtime=0) as compressed:
            shutil.copyfileobj(src, compressed)
    with gzip.open(target, 'rb') as stream:
        assert hashlib.file_digest(stream, 'sha256').hexdigest() == sha(path)
    record = dict(path=path.relative_to(ROOT).as_posix(), sha256=sha(path), bytes=path.stat().st_size,
                  compressed_path=target.relative_to(ROOT).as_posix(), compressed_sha256=sha(target), compressed_bytes=target.stat().st_size)
    path.unlink()
    return record

def export():
    check()
    records = []
    for lane in ('cpu-P1', 'cpu-P2'):
        out = ROOT / lane
        for name, argv in (
            ('perf-script', ['perf', 'script', '--no-inline', '-i', str(out / 'perf.data'), '-F', 'comm,pid,tid,time,event,period,ip,sym,dso']),
            ('top-symbols', ['perf', 'report', '--stdio', '--no-children', '--no-inline', '-g', 'none', '-i', str(out / 'perf.data'), '--sort', 'symbol', '--percent-limit', '0.5']),
        ):
            assert command(argv, out / name) == 0
        records += [compress(out / 'perf.data'), compress(out / 'perf-script.stdout')]
    for lane in ('heap-H1', 'heap-H2'):
        out = ROOT / lane
        heaps = [p for p in out.glob('heaptrack.*') if p.suffix in ('.zst', '.gz')]
        assert len(heaps) == 1, heaps
        heap = heaps[0]
        assert command(['heaptrack_print', '-f', str(heap), '-n', '30'], out / 'print') == 0
        assert command(['heaptrack_print', '-f', str(heap), '-t', '0', '-m', '0', '-p', '0', '-T', '0', '-l', '0',
                        '--filter-bt-function', 'litchi_perf_baseline::pptx_streaming_create::run', '-n', '40',
                        '-F', str(out / 'allocation-stacks.txt'), '--flamegraph-cost-type', 'allocations'], out / 'runner-allocations') == 0
        if heap.suffix == '.zst':
            assert command(['zstd', '--decompress', '--stdout', str(heap)], out / 'decoded') == 0
        else:
            assert command(['gzip', '--decompress', '--stdout', str(heap)], out / 'decoded') == 0
        records += [compress(out / 'decoded.stdout'), compress(out / 'print.stdout'), compress(out / 'runner-allocations.stdout'), compress(out / 'allocation-stacks.txt')]
    write(ROOT / 'compression.json', dict(schema='litchi-0475-compression-v1', artifacts=records))
    check()

if __name__ == '__main__':
    action = sys.argv[1]
    if action == 'prepare': prepare()
    elif action == 'export': export()
    elif action == 'matrix':
        for lane in json.loads((ROOT / 'protocol.json').read_text())['order']:
            capture(lane['lane'])
    else: capture(action)
