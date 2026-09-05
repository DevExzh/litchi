#!/usr/bin/env python3
"""Capture the frozen MCE candidate's residual CPU profile after ABBA."""
import datetime
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import time

out = Path('/tmp/litchi-goal-0410-profile')
out.mkdir(exist_ok=False)
wt = Path('/tmp/litchi-goal-0410-worktree')
bins = Path('/tmp/litchi-goal-0410-binaries/candidate')
identity = json.loads((bins / 'identity.json').read_text())
env = os.environ.copy()
env.update(DEBUGINFOD_URLS='', RUSTUP_TOOLCHAIN='1.98.1',
           RUSTFLAGS='-C force-frame-pointers=yes -C force-unwind-tables=yes')
commands = []
def run(label, argv):
    started = datetime.datetime.now(datetime.timezone.utc).isoformat()
    start = time.monotonic()
    with (out / (label + '.stdout')).open('wb') as stdout, (out / (label + '.stderr')).open('wb') as stderr:
        result = subprocess.run(list(map(str, argv)), cwd=wt, env=env, stdout=stdout, stderr=stderr)
    commands.append(dict(label=label, argv=list(map(str, argv)), cwd=str(wt),
                         started_utc=started, wall_seconds=time.monotonic()-start,
                         exit_code=result.returncode))
    (out / 'commands.json').write_text(json.dumps(commands, indent=2)+'\n')
    result.check_returncode()
def state():
    return dict(revision=subprocess.check_output(['git','rev-parse','HEAD'],cwd=wt,text=True).strip(),
                status=subprocess.check_output(['git','status','--porcelain=v1'],cwd=wt,text=True).strip())
assert state()['status'] == ''
run('checkout', ['git','checkout','--detach',identity['revision']])
before = state()
assert before == dict(revision=identity['revision'], status='')
binary = bins / 'litchi-perf-baseline'
assert hashlib.sha256(binary.read_bytes()).hexdigest() == identity['binaries'][binary.name]['sha256']
(out / 'environment.json').write_text(json.dumps(dict(source_before=before, binary_identity=identity,
    runtime_environment={k:env[k] for k in ['DEBUGINFOD_URLS','RUSTUP_TOOLCHAIN','RUSTFLAGS']},
    scope='whole process and inherited fresh children; selected ancestor attribution is sampled CPU, not phase latency'),indent=2)+'\n')
shutil.copy2(__file__, out / 'capture-profile.py')
shutil.copy2('/tmp/litchi-goal-0410-attribute.py', out / 'attribute_0410.py')
for kind in ['normal','allocator']:
    for suffix in ['.json','.catalog.json']:
        shutil.copy2(Path('/tmp/litchi-goal-0410-capture') / ('b1-'+kind+suffix), out / ('selected-'+kind+suffix))
run('selected-perf-record', ['taskset','-c','2','perf','record','--no-buildid-cache','-e','cycles:u','-F','999',
    '--call-graph','fp,127','-o',out/'perf.data','--',binary,'--case','xlsx_file_selected_cell',
    '--filesystem-cache','warm','--xlsx-cell-crud-shape','medium','--warmup','10','--samples','300',
    '--json',out/'selected-perf-record.json','--corpus-manifest',out/'selected-perf-record.catalog.json'])
assert state() == before
run('perf-all-self', ['perf','report','--stdio','--no-inline','--no-children','--call-graph=none','--percent-limit=0','-i',out/'perf.data'])
run('perf-script-no-inline', ['perf','script','--no-inline','-i',out/'perf.data'])
run('attribution', ['python3',out/'attribute_0410.py',
    '--script',out/'perf-script-no-inline.stdout','--report',out/'perf-all-self.stdout','--capture',out,
    '--repo','/home/zhuhe/code/litchi','--output',out/'attribution-summary.json'])
run('flamegraph', ['flamegraph','--perfdata',out/'perf.data','--no-inline','--deterministic','--title','XLSX selected cell MCE candidate: whole process and children','--output',out/'cpu-flamegraph.svg'])
print('Candidate residual profile captured and attributed')
