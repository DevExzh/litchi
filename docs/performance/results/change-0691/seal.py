#!/usr/bin/env python3
"""Seal final and archived bindings after cleanup, then recheck final docs."""
import hashlib
import json
import subprocess
from pathlib import Path

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
records = []
commands = [('audit', ['python3',str(P/'audit.py')]),
            ('audit-initial', ['python3',str(P/'audit.py'),'--initial']),
            ('production-unchanged', ['git','diff','HEAD','--exit-code','--','crates','Cargo.toml'])]
for row in json.loads((P/'evidence/results.json').read_text()):
    if row['name'] in ['report','coverage','non-iwork','claims-structural']:
        commands.append(('final-'+row['name'],row['command']))
for name,command in commands:
    output = P / 'evidence' / (name+'.log')
    with output.open('w') as log:
        result = subprocess.run(command,cwd=ROOT,stdout=log,stderr=subprocess.STDOUT)
    records.append(dict(name=name,command=command,exit_code=result.returncode,
                        log_sha256=hashlib.sha256(output.read_bytes()).hexdigest()))
    (P/'final-validation.json').write_text(json.dumps(records,indent=2)+'\n')
    print(name,result.returncode,flush=True)
    assert result.returncode==0
boundary = ROOT/'docs/performance/results/change-0690/evidence/crate-boundaries.log'
(P/'boundary-reuse.json').write_text(json.dumps(dict(
    evidence=str(boundary.relative_to(ROOT)),sha256=hashlib.sha256(boundary.read_bytes()).hexdigest(),
    reason='No workspace or crate source/manifest changed from the previously checked baseline; the new standalone probe is outside the workspace.'
),indent=2)+'\n')
