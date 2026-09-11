#!/usr/bin/env python3
"""Capture serial AB/BA export timings or one instrumented after profile."""
import datetime
import json
import shutil
import subprocess
import sys
import time
from pathlib import Path
from run import HERE, REPO, SCRATCH, BINARY, sha, sources, write

def run(stage, name, kind='timing', shapes=None, samples=None):
    dest = HERE / stage
    dest.mkdir(exist_ok=True)
    binary = SCRATCH / stage / 'litchi-perf-baseline'
    source_dir = HERE / 'before' if stage == 'before' else HERE
    receipt = json.loads((source_dir / 'build-receipt.json').read_text())
    assert sha(binary) == receipt['binary_sha256']
    final_sources = sources()
    current_manifest = HERE / 'source-manifest.json'
    if not current_manifest.exists():
        current_manifest = HERE / 'before/source-manifest.json'
    assert final_sources == json.loads(current_manifest.read_text())
    command = [str(binary), '--case', 'odt_semantic_text_to_sink', '--semantic-shape', shapes or ('tiny,medium,large' if kind == 'timing' else 'large'), '--warmup', '10' if kind == 'timing' else '0', '--samples', str(samples) if samples else ('500' if kind == 'timing' else ('1000' if kind in {'perf-stat','perf-record'} else ('20' if kind == 'heaptrack' else '5'))), '--json', str(dest / f'{name}-report.json'), '--corpus-manifest', str(dest / f'{name}-catalog.json')]
    if kind == 'timing':
        command = ['taskset', '-c', '2', *command]
    if kind == 'heaptrack':
        command = ['heaptrack', '-o', str(SCRATCH / stage / 'heaptrack'), *command]
    elif kind == 'callgrind':
        command = ['valgrind', '--tool=callgrind', '--collect-atstart=no', '--toggle-collect=*write_text_blocks_to_writer*', '--callgrind-out-file=' + str(dest / 'callgrind.out'), *command]
    if kind == 'perf-stat':
        command = ['taskset','-c','2','perf','stat','-x',',','-o',str(dest / f'{name}.csv'),'-e','{cycles,instructions,branches,branch-misses},page-faults,context-switches,cpu-migrations','--',*command]
    elif kind == 'perf-record':
        command = ['taskset','-c','2','perf','record','--no-buildid-cache','-F','499','--call-graph','dwarf,8192','-o',str(dest / 'perf.data'),'--',*command]
    load_before = Path('/proc/loadavg').read_text().strip()
    started = datetime.datetime.now(datetime.timezone.utc).isoformat()
    tick = time.monotonic()
    with (dest / f'{name}.log').open('x') as log:
        child = subprocess.run(['/usr/bin/time', '-v', *command], cwd=REPO, stdout=log, stderr=subprocess.STDOUT)
    assert child.returncode == 0
    assert sources() == final_sources and sha(binary) == receipt['binary_sha256']
    write(dest / f'{name}-receipt.json', {'command': command, 'started_utc': started, 'host_loadavg_before': load_before, 'host_loadavg_after': Path('/proc/loadavg').read_text().strip(), 'elapsed_seconds': time.monotonic()-tick, 'exit_code': child.returncode, 'binary_sha256': sha(binary), 'build_receipt_sha256': sha(source_dir / 'build-receipt.json'), 'source_manifest_sha256': sha(source_dir / 'source-manifest.json'), 'current_source_unchanged': True, 'scope': 'export-only uninstrumented wall-clock; hashing inside timer' if kind == 'timing' else 'instrumented, excluded from latency comparison', 'artifacts': {p.name: sha(p) for p in [dest / f'{name}.log', dest / f'{name}-report.json', dest / f'{name}-catalog.json']}})
    if kind == 'heaptrack':
        shutil.copy2(SCRATCH / stage / 'heaptrack.zst', dest / 'heaptrack.zst')
    print(stage, name, child.returncode, flush=True)

if __name__ == '__main__':
    if sys.argv[1].startswith('before-'):
        kind = sys.argv[1].removeprefix('before-')
        run('before',kind,kind)
    elif sys.argv[1] == 'hardware':
        for stage,name in [('before','perf-stat-r1'),('after','perf-stat-r1'),('after','perf-stat-r2'),('before','perf-stat-r2')]:
            run(stage,name,'perf-stat')
    elif sys.argv[1] == 'tail-followup':
        for stage,name in [('before','tail-r1'),('after','tail-r1'),('after','tail-r2'),('before','tail-r2')]:
            run(stage,name,shapes='large',samples=2000)
    elif sys.argv[1] == 'timing':
        for stage, name in [('before','r1'), ('after','r1'), ('after','r2'), ('before','r2')]:
            run(stage, name)
    else:
        run('after', sys.argv[1], sys.argv[1])
