#!/usr/bin/env python3
"""Recheck the retained packet and final documentation after scratch cleanup."""
import hashlib,json,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
commands=[('audit',['python3',str(P/'audit.py')])]
commands += [('final-'+r['name'],r['command']) for r in json.loads((P/'gates.json').read_text()) if r['name'] in ['claims-structural','report','coverage','non-iwork']]
assert len(commands)==5
records=[]
for name,command in commands:
 log=P/'gates'/(name+'.log')
 with log.open('w') as output:
  result=subprocess.run(command,cwd=ROOT,stdout=output,stderr=subprocess.STDOUT)
 records.append(dict(name=name,command=command,exit_code=result.returncode,log_sha256=hashlib.sha256(log.read_bytes()).hexdigest()))
 (P/'final-validation.json').write_text(json.dumps(records,indent=2)+'\n')
 print(name,result.returncode,flush=True)
 assert result.returncode==0,log.read_text()[-3000:]
