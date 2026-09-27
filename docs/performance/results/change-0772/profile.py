"""Qualify ordinary-binary stacks at the exact timed OPC publication call.

Run only after other root-owned Cargo/native work has completed. This capture
does not itself assert that stack coverage is adequate for attribution.
"""
import gzip
import json
from pathlib import Path
import shutil
import subprocess
import time

from build import P, ROOT, guard, sha


if __name__ == '__main__':
    guard()
    build = json.loads((P / 'build.json').read_text())
    binary = Path(build['binary'])
    assert sha(binary) == build['binary_sha256']
    folder = P / 'profile-qualification'
    folder.mkdir()
    report = folder / 'report.json'
    data = folder / 'perf.data'
    native = ['taskset', '-c', '12', str(binary), '--warmup', '0', '--samples', '30',
              '--case', 'opc_mutated_save', '--shape', 'few-large', '--payload', 'incompressible',
              '--json', str(report)]
    command = ['perf', 'record', '--no-buildid-cache', '-e', 'cycles:u', '-F', '199', '--call-graph', 'dwarf,32768',
               '-o', str(data), '--', *native]
    row = {'command': command, 'started': time.time(), 'monotonic_start': time.monotonic_ns(),
           'binary_sha256': build['binary_sha256'], 'boundary_sha256': sha(P / 'timed-boundary.json')}
    with (folder / 'stdout').open('w') as out, (folder / 'stderr').open('w') as err:
        result = subprocess.run(command, cwd=ROOT, stdout=out, stderr=err)
    row |= {'exit': result.returncode, 'ended': time.time(), 'monotonic_end': time.monotonic_ns()}
    (folder / 'manifest.json').write_text(json.dumps(row, indent=2) + '\n')
    assert result.returncode == 0, row
    command = ['perf', 'script', '--sym-offset', '--show-lost-events', '-i', str(data),
               '-F', 'comm,pid,tid,time,event,period,ip,sym,dso']
    with (folder / 'script.stderr').open('w') as err, (folder / 'stacks.txt.gz').open('wb') as raw:
        with gzip.GzipFile(filename='', mode='wb', fileobj=raw, mtime=0) as zipped:
            process = subprocess.Popen(command, cwd=ROOT, stdout=subprocess.PIPE, stderr=err)
            shutil.copyfileobj(process.stdout, zipped)
            code = process.wait()
    row |= {'script_command': command, 'script_exit': code}
    row['files'] = {str(f.relative_to(P)): sha(f) for f in folder.iterdir() if f.is_file() and f.name != 'manifest.json'}
    (folder / 'manifest.json').write_text(json.dumps(row, indent=2) + '\n')
    assert code == 0
    guard()
    assert sha(binary) == build['binary_sha256']
    print('Captured ordinary-binary profile; stack qualification remains required', flush=True)
