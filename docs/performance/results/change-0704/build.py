#!/usr/bin/env python3
"""Freeze the 0704 native/allocation binaries around the PPTX memo candidate.

The baseline is exact against the recorded 602-file census.  A candidate may
add or change PPTX Rust sources (the memo module and its tests); it must leave
the shared OOXML and OPC Rust sources byte-identical.  This keeps the packet
from silently attributing a shared-codec edit to the PPTX experiment.
"""
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
target = ROOT.parent/'litchi-target-0704'
bin_dir = ROOT.parent/'litchi-0704-bin'
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
if phase == 'baseline':
    assert source == base['source_sha256']
else:
    missing = set(base['source_sha256']) - set(source)
    assert not missing, f'candidate removed baseline source files: {sorted(missing)[:5]}'
    shared = [name for name, digest in base['source_sha256'].items()
              if name.startswith(('crates/litchi-ooxml-common/', 'crates/litchi-opc/'))
              and source.get(name) != digest]
    assert not shared, f'candidate changed shared source outside PPTX lane: {shared[:5]}'
    changed = [name for name in base['source_sha256']
               if source.get(name) != base['source_sha256'][name]]
    added = sorted(set(source) - set(base['source_sha256']))
    assert all(name.startswith('crates/litchi-pptx/') for name in changed + added), \
        'candidate source census contains a non-PPTX change'
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
    shutil.copy2(target/'release/probe0704',output)
    records.append(dict(label=label,phase=phase,command=command,exit_code=result.returncode,
                        seconds=time.monotonic()-started,binary=str(output),binary_sha256=sha(output),
                        probe_sha256={str(f.relative_to(P)):sha(f) for f in (P/'probe').rglob('*') if f.is_file()},
                        source_sha256=source,build_inputs_sha256=base['build_inputs_sha256'],rustflags='-D warnings'))
    (P/f'builds-{phase}.json').write_text(json.dumps(records,indent=2)+'\n')
    print(phase,label,'built',flush=True)
