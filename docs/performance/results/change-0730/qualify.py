#!/usr/bin/env python3
"""Qualify final candidate against exact 0728 DOC output and preservation gates."""
import hashlib,json,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
def read(p):return json.loads(p.read_text())
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def write(p,x):p.write_text(json.dumps(x,indent=2)+'\n')
subprocess.run(['python3',str(P/'source-guard.py')],cwd=ROOT,check=True)
i=0
while (P/f'qualification-{i}').exists():i+=1
out=P/f'qualification-{i}';out.mkdir();b=read(P/'candidate-builds.json');binary=next(r for r in b['binaries'] if Path(r['path']).name=='candidate-ole_format_save_probe');assert sha(Path(binary['path']))==binary['sha256'];contracts=read(P/'oracle-contract.json');runs=[];smoke=[]
for c in read(P/'cases.json'):
 name=c['case']+'.json';cmd=['taskset','-c','12',binary['path'],'--case',c['case'],'--input',c['path'],'--operation','format','--samples','1','--warmups','1']
 with (out/name).open('wb') as f,(out/(name+'.stderr')).open('wb') as e:r=subprocess.run(cmd,cwd=ROOT,stdout=f,stderr=e)
 runs.append(dict(command=cmd,exit_code=r.returncode,output=name,sha256=sha(out/name),stderr_sha256=sha(out/(name+'.stderr'))));write(out/'manifest.json',dict(runs=runs,candidate_builds_sha256=sha(P/'candidate-builds.json')));assert r.returncode==0
 x=read(out/name);cc=contracts[c['case']];assert {k:x[k] for k in cc['identity']}==cc['identity'];assert x['expected_oracle']['semantic_witness']==cc['semantic_witness'];assert x['allocator_instrumented'] is False and x['timing_claim'] is True
 assert [o['name'] for o in x['oracle_controls']]==cc['control_names'] and all(o['rejected'] and o['status']=='rejected' and o['failure_reasons'] for o in x['oracle_controls'])
 assert len(x['samples'])==1
 for o in [x['expected_oracle']]+[s['oracle'] for s in x['samples']]:assert all(v for v in o.values() if isinstance(v,bool)) and not o['failure_reasons']
 s=x['samples'][0];assert s['output_sha256']==x['expected_output_sha256'];assert s['oracle']['semantic_witness']==cc['semantic_witness']
 smoke.append(dict(case=c['case'],format='doc',input=c['path'],source_sha256=c['sha256'],expected_output_sha256=x['expected_output_sha256'],output_sha256=s['output_sha256'],exit_code=0,oracle_ok=True))
write(out/'candidate-smoke.json',dict(status='pass',kind='candidate-smoke',candidate_builds_sha256=sha(P/'candidate-builds.json'),retention_builds_sha256=sha(P/'retention-builds.json'),runs=smoke))
quality=max((q for q in P.glob('quality-*') if q.is_dir()),key=lambda q:int(q.name.split('-')[1]));m=read(quality/'manifest.json');assert len(m['runs'])==13 and all(r['exit_code']==0 for r in m['runs']);assert m['source_sha256']=={k:v for k,v in b['source_sha256'].items() if k.startswith('crates/')}
files={str(p.relative_to(P)):sha(p) for d in [quality,out] for p in d.iterdir() if p.is_file()}
write(P/'qualification.json',dict(status='pass',files=files,candidate_builds_sha256=sha(P/'candidate-builds.json'),retention_builds_sha256=sha(P/'retention-builds.json')));print('PASS candidate exact 0728 output/oracle; 13 final quality commands bound')
