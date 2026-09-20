#!/usr/bin/env python3
"""Seal retained bindings and final documentation after owned scratch cleanup."""
import hashlib,json,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
commands=[('audit',['python3',str(P/'audit.py')])]
for row in json.loads((P/'evidence/results.json').read_text()):
    if row['name'] in ['report','coverage','non-iwork','claims-structural']:
        commands.append(('final-'+row['name'],row['command']))
records=[]
for name,command in commands:
    path=P/'evidence'/(name+'.log')
    with path.open('w') as log:
        result=subprocess.run(command,cwd=ROOT,stdout=log,stderr=subprocess.STDOUT)
    records.append(dict(name=name,command=command,exit_code=result.returncode,log_sha256=hashlib.sha256(path.read_bytes()).hexdigest()))
    (P/'final-validation.json').write_text(json.dumps(records,indent=2)+'\n')
    print(name,result.returncode,flush=True)
    assert result.returncode==0,path.read_text()
