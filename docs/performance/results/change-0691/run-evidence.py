#!/usr/bin/env python3
"""Recheck documentation/claim gates; production boundary inputs are unchanged."""
import hashlib
import json
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
    if name == 'crate-boundaries':
        continue
    started = time.monotonic()
    with (out / (name+'.log')).open('w') as log:
        result = subprocess.run(previous['command'],cwd=ROOT,stdout=log,stderr=subprocess.STDOUT)
    records.append(dict(name=name,command=previous['command'],exit_code=result.returncode,
                        seconds=time.monotonic()-started,
                        log_sha256=hashlib.sha256((out / (name+'.log')).read_bytes()).hexdigest()))
    (out / 'results.json').write_text(json.dumps(records,indent=2)+'\n')
    print(name,result.returncode,flush=True)
    assert result.returncode==0
