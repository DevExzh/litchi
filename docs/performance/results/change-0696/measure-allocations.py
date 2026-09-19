#!/usr/bin/env python3
"""Separate allocator diagnostics; no timings from these runs are native evidence."""
import hashlib
import json
import subprocess
from pathlib import Path

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
import sys
phase=sys.argv[1]
assert phase in ['baseline','candidate']
binary = ROOT.parent / 'litchi-0696-bin' / (phase+'-allocations')
native = json.loads((P / 'native-runs-baseline.json').read_text())
cases = {r['case']:(r['command'][5],r['workflow']) for r in native}
out = P / 'allocations'
out.mkdir(exist_ok=True)
runs = []
for case, (source, workflow) in cases.items():
    command = ['taskset','-c','12',str(binary),'phases',source,'3','2',workflow]
    output = out / (case+'-'+phase+'.tsv')
    with output.open('w') as stdout, (out / (case+'-'+phase+'.stderr')).open('w') as stderr:
        result = subprocess.run(command,cwd=ROOT,stdout=stdout,stderr=stderr)
    runs.append(dict(case=case,phase=phase,workflow=workflow,command=command,exit_code=result.returncode,
                     output=str(output.relative_to(P)),
                     output_sha256=hashlib.sha256(output.read_bytes()).hexdigest(),
                     binary_sha256=hashlib.sha256(binary.read_bytes()).hexdigest()))
    (P / ('allocation-runs-'+phase+'.json')).write_text(json.dumps(runs,indent=2)+'\n')
    assert result.returncode==0,case
    print(case,'allocations captured',flush=True)
