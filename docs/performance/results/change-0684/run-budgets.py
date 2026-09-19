#!/usr/bin/env python3
"""Candidate-only cache-limit controls; timing is diagnostic, counters primary."""
import json,hashlib,subprocess,sys
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3];binary=sys.argv[1];D=P/'budgets';D.mkdir(exist_ok=True)
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
commands=[]
for case in json.loads((P/'cases.json').read_text()):
 if not any(case['case'].startswith(v) for v in ['54016-','WithCustomViews-','Simple-']):continue
 reference=None
 for budget in [0,1,1048576,2097152]:
  command=['taskset','-c','12',binary,'--input',case['path'],'--worksheet',str(case['sheet']),'--row',str(case['row']),'--column',str(case['column']),'--budget',str(budget),'--queries','5']
  r=subprocess.run(command,cwd=ROOT,capture_output=True,text=True);name=f'{case["case"]}-{budget}.json';(D/name).write_text(r.stdout)
  commands.append(dict(command=command,exit_code=r.returncode,output=name));assert r.returncode==0,r.stderr
  value=json.loads(r.stdout)
  print(case['case'],budget,flush=True)
(D/'commands.json').write_text(json.dumps(commands,indent=2)+'\n')
m=dict(binary_sha256=sha(binary),source_sha256={str(p.relative_to(ROOT)):sha(p) for p in sorted((ROOT/'crates/litchi-xls').rglob('*.rs'))},probe_sha256={str(p.relative_to(P)):sha(p) for p in sorted((P/'budget-probe').rglob('*')) if p.is_file()},cases_sha256=sha(P/'cases.json'),raw_sha256={p.name:sha(p) for p in D.iterdir() if p.is_file() and p.name!='manifest.json'})
(D/'manifest.json').write_text(json.dumps(m,indent=2)+'\n')
