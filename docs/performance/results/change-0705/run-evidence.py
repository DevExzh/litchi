#!/usr/bin/env python3
"""Recheck all boundary, documentation and claim gates."""
import hashlib
import json
import os
import subprocess
import time
from pathlib import Path

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
out = P / 'evidence'
out.mkdir(exist_ok=True)
records = []
for previous in json.loads((P.parent / 'change-0690/evidence/results.json').read_text()):
    name = previous['name']
    head = subprocess.check_output(['git','rev-parse','HEAD'],cwd=ROOT,text=True).strip()
    source_sha256 = {str(f.relative_to(ROOT)):hashlib.sha256(f.read_bytes()).hexdigest() for c in ['litchi-xlsx','litchi-ooxml-common','litchi-opc'] for f in (ROOT/'crates'/c).rglob('*.rs')}
    started = time.monotonic()
    with (out / (name+'.log')).open('w') as log:
        result = subprocess.run(previous['command'],cwd=ROOT,env={**os.environ,"CARGO_BUILD_JOBS":"2","CARGO_TARGET_DIR":str(ROOT.parent/"litchi-target-0705")},stdout=log,stderr=subprocess.STDOUT)
    records.append(dict(name=name,head=head,source_sha256=source_sha256,command=previous['command'],exit_code=result.returncode,
                        seconds=time.monotonic()-started,
                        log_sha256=hashlib.sha256((out / (name+'.log')).read_bytes()).hexdigest()))
    (out / 'results.json').write_text(json.dumps(records,indent=2)+'\n')
    print(name,result.returncode,flush=True)
    assert result.returncode==0
