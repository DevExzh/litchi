#!/usr/bin/env python3
"""Profile the native commit prefix; raw perf data stays in owned scratch.

The ``commit`` prefix includes package open, presentation capture, working
clone, one text edit, and commit.  It therefore includes the unchanged-slide
recapture work that the 0704 candidate is intended to remove.  It stops before
publication and serialization, so its denominator is distinct from the
end-to-end native phase timer.
"""
import hashlib
import json
import subprocess
import time
from pathlib import Path

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
import sys
phase=sys.argv[1]
assert phase in ['baseline','candidate']
binary = ROOT.parent / 'litchi-0704-bin' / (phase+'-native')
scratch = ROOT.parent / 'litchi-0704-profile'
scratch.mkdir(exist_ok=True)
out = P / 'profile' / phase
out.mkdir(parents=True, exist_ok=True)
source = ROOT / 'test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-section-test.pptx'
commands = []
def run(name, command):
    started = time.monotonic()
    with (out / (name+'.txt')).open('w') as log:
        result = subprocess.run(command, cwd=ROOT, stdout=log, stderr=subprocess.STDOUT)
    commands.append(dict(name=name, command=command, exit_code=result.returncode,
                         seconds=time.monotonic()-started))
    (out / 'commands.json').write_text(json.dumps(commands, indent=2)+'\n')
    print(name, result.returncode, flush=True)
    return result.returncode

data = scratch / (phase+'-commit.data')
rc = run('record', ['perf', 'record', '-F', '997', '-g', '--call-graph', 'dwarf,16384',
                     '-o', str(data), '--', 'taskset', '-c', '12', str(binary),
                     'prefix', str(source), 'commit', '800'])
if rc == 0:
    run('self-symbols', ['perf', 'report', '--stdio', '--no-children', '--percent-limit', '0.1', '-i', str(data)])
    run('inclusive-symbols', ['perf', 'report', '--stdio', '--children', '--percent-limit', '1', '-i', str(data)])
for n in [10, 210]:
    run(f'counters-{n}', ['perf', 'stat', '-x', '\t', '-e',
                          'cycles,instructions,branches,branch-misses,cache-misses,page-faults,task-clock',
                          '--', 'taskset', '-c', '12', str(binary), 'prefix', str(source), 'commit', str(n)])
run('rss', ['/usr/bin/time', '-v', 'taskset', '-c', '12', str(binary), 'prefix', str(source), 'commit', '210'])
(out / 'binding.json').write_text(json.dumps(dict(
    binary_sha256=hashlib.sha256(binary.read_bytes()).hexdigest(),
    source_sha256=hashlib.sha256(source.read_bytes()).hexdigest(),
    raw_data_sha256=hashlib.sha256(data.read_bytes()).hexdigest() if data.exists() else None,
    files={f.name:hashlib.sha256(f.read_bytes()).hexdigest() for f in out.iterdir() if f.is_file() and f.name!='binding.json'}
), indent=2)+'\n')
