#!/usr/bin/env python3
"""Freeze native/allocation binaries before and after the shared MCE change."""
import hashlib
import json
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
phase = sys.argv[1]
assert phase in ['baseline','candidate']
target = ROOT.parent/'litchi-target-0700'
bin_dir = ROOT.parent/'litchi-0700-bin'
bin_dir.mkdir(exist_ok=True)
def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()
base = json.loads((P/'baseline.json').read_text())
for name,digest in base['constraints_sha256'].items():
    assert sha(ROOT/name)==digest,name
for name,digest in base['build_inputs_sha256'].items():
    assert sha(ROOT/name)==digest,name
probe_inputs={str(f.relative_to(P)):sha(f) for f in (P/'probe').rglob('*') if f.is_file()}
if phase=='candidate':
    for previous in json.loads((P/'builds-baseline.json').read_text()):
        assert previous['probe_sha256']==probe_inputs
source = {str(f.relative_to(ROOT)):sha(f) for c in ['litchi-pptx','litchi-ooxml-common','litchi-opc']
          for f in (ROOT/'crates'/c).rglob('*.rs')}
if phase=='baseline':
    assert source==base['source_sha256']
records=[]
for label,features in [('native',[]),('allocations',['--features','allocations'])]:
    command=['cargo','build','--release','--locked','--manifest-path',str(P/'probe/Cargo.toml'),
             '--target-dir',str(target),'-j','2',*features]
    started=time.monotonic()
    with (P/f'build-{phase}-{label}.log').open('w') as log:
        result=subprocess.run(command,cwd=ROOT,env={**os.environ,'RUSTFLAGS':'-D warnings'},
                              stdout=log,stderr=subprocess.STDOUT)
    assert result.returncode==0,label
    assert source == {str(f.relative_to(ROOT)):sha(f) for c in ['litchi-pptx','litchi-ooxml-common','litchi-opc'] for f in (ROOT/'crates'/c).rglob('*.rs')}, 'source changed during build'
    assert probe_inputs == {str(f.relative_to(P)):sha(f) for f in (P/'probe').rglob('*') if f.is_file()}
    for name,digest in base['build_inputs_sha256'].items():
        assert sha(ROOT/name)==digest,name
    output=bin_dir/f'{phase}-{label}'
    shutil.copy2(target/'release/probe0700',output)
    records.append(dict(label=label,phase=phase,command=command,exit_code=result.returncode,
                        seconds=time.monotonic()-started,binary=str(output),binary_sha256=sha(output),
                        probe_sha256={str(f.relative_to(P)):sha(f) for f in (P/'probe').rglob('*') if f.is_file()},
                        source_sha256=source,build_inputs_sha256=base['build_inputs_sha256'],rustflags='-D warnings'))
    (P/f'builds-{phase}.json').write_text(json.dumps(records,indent=2)+'\n')
    print(phase,label,'built',flush=True)
