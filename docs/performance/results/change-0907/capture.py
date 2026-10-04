"""Serial, fixed-CPU same-binary PPTX requalification. Refuses existing outputs."""
import datetime
import hashlib
import json
import os
from pathlib import Path
import platform
import subprocess

P = Path(__file__).resolve().parent
S = Path('/tmp/litchi-0907-src')
B = Path('/home/zhuhe/code/litchi/target/litchi-0907-build/release/litchi-perf-baseline')
ENV = {'PATH': '/usr/bin:/bin', 'LC_ALL': 'C', 'LANG': 'C', 'TZ': 'UTC'}
CPU = 12
LEGS = [('eager-a', 'eager'), ('source-a', 'source_backed'), ('source-b', 'source_backed'), ('eager-b', 'eager')]

def now():
    return datetime.datetime.now(datetime.timezone.utc).isoformat()

def desc(path):
    return {'path': str(path), 'bytes': path.stat().st_size, 'sha256': hashlib.sha256(path.read_bytes()).hexdigest()}

def write(path, value):
    with path.open('x') as out:
        json.dump(value, out, indent=2, sort_keys=True)
        out.write('\n')

def source_check():
    manifest = json.loads((P / 'source.json').read_text())
    for row in manifest['files']:
        actual = desc(S / row['path'])
        assert actual['bytes'] == row['bytes'] and actual['sha256'] == row['sha256'], row['path']
    return {'file_count': len(manifest['files']), 'manifest': desc(P / 'source.json'), 'checked_utc': now()}

assert CPU in os.sched_getaffinity(0)
assert json.loads((P / 'native-build-v2-run.json').read_text())['exit_code'] == 0
for label, _ in LEGS:
    for suffix in ['.json', '.stdout', '.stderr', '-start.json', '-run.json']:
        assert not (P / (label + suffix)).exists(), label + suffix
binary = desc(B)
write(P / 'binary.json', binary)
write(P / 'environment.json', {'recorded_utc': now(), 'platform': platform.uname()._asdict(), 'cpu': CPU, 'affinity': sorted(os.sched_getaffinity(0)), 'lscpu': subprocess.check_output(['lscpu'], text=True), 'rustc': subprocess.check_output(['/home/zhuhe/.cargo/bin/rustc', '+1.95.0', '-Vv'], text=True), 'meminfo': Path('/proc/meminfo').read_text(), 'scope': 'Shared host; no exclusive CPU or host claim. Logical in-memory reads, not device I/O.'})
write(P / 'capture-plan.json', {'legs': LEGS, 'warmup': 30, 'samples_per_leg': 200, 'cpu': CPU, 'environment': ENV, 'binary': binary, 'source_before': source_check(), 'helper': desc(Path(__file__)), 'scope': 'Descriptive same-binary guarded workflow comparison; no candidate optimization or historical timing comparison.'})
for label, kind in LEGS:
    assert desc(B) == binary
    argv = ['taskset', '-c', str(CPU), str(B), '--case', 'pptx_' + kind + '_multi_slide_batch_edit_save', '--workers', '1', '--warmup', '30', '--samples', '200', '--json', str(P / (label + '.json'))]
    receipt = {'argv': argv, 'cwd': str(S), 'environment': ENV, 'started_utc': now(), 'load_before': os.getloadavg(), 'binary': binary}
    write(P / (label + '-start.json'), receipt)
    print('start', label, receipt['started_utc'], flush=True)
    with (P / (label + '.stdout')).open('x') as stdout, (P / (label + '.stderr')).open('x') as stderr:
        result = subprocess.run(argv, cwd=S, env=ENV, stdout=stdout, stderr=stderr)
    receipt.update(exit_code=result.returncode, finished_utc=now(), load_after=os.getloadavg(), stdout=desc(P / (label + '.stdout')), stderr=desc(P / (label + '.stderr')))
    if (P / (label + '.json')).exists():
        receipt['result'] = desc(P / (label + '.json'))
    write(P / (label + '-run.json'), receipt)
    assert result.returncode == 0, receipt
    print('finish', label, receipt['finished_utc'], flush=True)
write(P / 'capture-complete.json', {'status': 'pass', 'completed_utc': now(), 'binary_unchanged': desc(B) == binary, 'source_after': source_check(), 'workload_children': len(LEGS), 'retained_samples': 800})
