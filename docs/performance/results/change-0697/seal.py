#!/usr/bin/env python3
"""Audit cleanup and recheck final documentation without repeating compilation."""
import hashlib,json,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
commands=[('audit',['python3',str(P/'audit.py')])]
commands += [('final-'+r['name'],r['command']) for r in json.loads((P/'validation.json').read_text()) if r['name'] in ['claims-structural','report','coverage','non-iwork']]
rows=[]
for name,command in commands:
 path=P/'validation'/(name+'.log')
 with path.open('w') as out:r=subprocess.run(command,cwd=ROOT,stdout=out,stderr=subprocess.STDOUT)
 rows.append(dict(name=name,command=command,exit_code=r.returncode,log_sha256=hashlib.sha256(path.read_bytes()).hexdigest()))
 (P/'final-validation.json').write_text(json.dumps(rows,indent=2)+'\n')
 print(name,r.returncode,flush=True)
 assert r.returncode==0,path.read_text()
