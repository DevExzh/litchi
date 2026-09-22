#!/usr/bin/env python3
"""Synthetic analyzer/schema integration; temporary records are not measurements."""
import contextlib,copy,hashlib,importlib.util,io,json,shutil,tempfile
from pathlib import Path
P=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('attribution',P/'analyze.py');mod=importlib.util.module_from_spec(spec);spec.loader.exec_module(mod)
def read(p):return json.loads(p.read_text())
def write(p,x):p.write_text(json.dumps(x,indent=2)+'\n')
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
q=max((p for p in P.glob('qualification-*') if p.is_dir()),key=lambda p:int(p.name.split('-')[-1]));b=read(P/'builds.json');cases={c['case']:c for c in read(P/'cases.json')};plan=read(P/'plan.json')
with tempfile.TemporaryDirectory(prefix='litchi-0729-preflight-') as temp:
 t=Path(temp);(t/'captures').mkdir()
 for name in ['plan.json','cases.json','oracle-contract.json']:shutil.copy2(P/name,t/name)
 write(t/'freeze.json',dict(bindings={},binaries=b['binaries']));m=dict(status='complete',freeze_sha256=sha(t/'freeze.json'),bindings_start={},bindings_end={},runs=[])
 for row in plan['schedule']:
  c=cases[row['case']];template=read(q/f"{row['case']}-{row['route']}.json");x=copy.deepcopy(template);x['samples_requested']=plan['samples'];x['warmups']=plan['warmups'];x['samples']=[dict(copy.deepcopy(template['samples'][0]),index=i) for i in range(plan['samples'])]
  out=f"c{row['cycle']}-{row['case']}-{row['route']}-r{row['repeat']}.json";write(t/'captures'/out,x);(t/'captures'/(out+'.stderr')).write_text('')
  cmd=['taskset','-c',str(plan['cpu']),b['binaries'][0]['path'],'--case',row['case'],'--input',c['path'],'--route',row['route'],'--samples',str(plan['samples']),'--warmups',str(plan['warmups'])]
  m['runs'].append(dict(row,exit_code=0,command=cmd,output=out,sha256=sha(t/'captures'/out),stderr=out+'.stderr',stderr_sha256=sha(t/'captures'/(out+'.stderr'))))
 write(t/'captures/manifest.json',m);mod.P=t
 with contextlib.redirect_stdout(io.StringIO()):mod.main()
 assert len(read(t/'analysis.json')['processes'])==72
 audit_spec=importlib.util.spec_from_file_location('independent',P/'audit.py');audit=importlib.util.module_from_spec(audit_spec);audit_spec.loader.exec_module(audit);audit.PACKET=t;audit.CAPTURES=t/'captures'
 audit.audit_capture(plan,cases,read(t/'oracle-contract.json'),read(t/'freeze.json'))
write(P/'preflight.json',dict(status='passed',kind='synthetic schema integration only; not measurement',analyzer_sha256=sha(P/'analyze.py'),auditor_sha256=sha(P/'audit.py'),contract_sha256=sha(P/'oracle-contract.json'),script_sha256=sha(Path(__file__))))
print('PASS synthetic schema/matrix integration; no timing evidence created')
