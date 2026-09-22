#!/usr/bin/env python3
"""Preserve all pre-freeze route qualification receipts, including failures."""
import hashlib,json,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def read(p):return json.loads(p.read_text())
def write(p,x):p.write_text(json.dumps(x,indent=2)+'\n')
i=0
while (P/f'qualification-{i}').exists():i+=1
out=P/f'qualification-{i}';out.mkdir();build=read(P/'builds.json');rows=[];identities={}
for b in build['binaries']:
 assert sha(Path(b['path']))==b['sha256']
 for c in read(P/'cases.json'):
  assert sha(ROOT/c['path'])==c['sha256'] and (ROOT/c['path']).stat().st_size==c['bytes']
  for op,policy in [('format','reuse'),('container','reuse'),('container','rewrite')]:
   name=f"{Path(b['path']).name}-{c['case']}-{op}-{policy}"
   cmd=['taskset','-c','12',b['path'],'--case',c['case'],'--input',c['path'],'--operation',op,'--policy',policy,'--samples','1','--warmups','0']
   with (out/(name+'.json')).open('wb') as stdout,(out/(name+'.stderr')).open('wb') as stderr:
    r=subprocess.run(cmd,cwd=ROOT,stdout=stdout,stderr=stderr)
   row=dict(command=cmd,exit_code=r.returncode,output=name+'.json',sha256=sha(out/(name+'.json')),stderr=name+'.stderr',stderr_sha256=sha(out/(name+'.stderr')));rows.append(row)
   write(out/'manifest.json',dict(builds_sha256=sha(P/'builds.json'),script_sha256=sha(Path(__file__)),runs=rows))
   if r.returncode==0:
    x=read(out/(name+'.json'));assert x['source_sha256']==c['sha256'] and x['source_inventory']['file_bytes']==c['bytes']
    assert (x['case'],x['operation'],x['policy'],x['samples_requested'],x['warmups'])==(c['case'],op,policy,1,0)
    identity={k:x[k] for k in ['source_inventory','expected_output_inventory','expected_output_sha256','replacements_sha256','replacements','changed_length_proof']}
    assert identity==identities.setdefault(c['case'],identity)
    assert x['changed_length_proof']['logical_stream_length_change_proven'] and x['changed_length_proof']['format_specific_semantic_length_proven']
    assert x['oracle_controls'] and all(c['rejected'] and c['status']=='rejected' and c['failure_reasons'] for c in x['oracle_controls'])
    assert all(v for v in x['expected_oracle'].values() if isinstance(v,bool)) and not x['expected_oracle']['failure_reasons']
    assert len(x['samples'])==1
    for sample in x['samples']:
     assert all(v for v in sample['oracle'].values() if isinstance(v,bool)) and not sample['oracle']['failure_reasons']
     assert sample['output_inventory']['streams']==x['expected_output_inventory']['streams']
   print(name,r.returncode,flush=True)
assert all(r['exit_code']==0 for r in rows)
print('PASS 18 separate route/lane qualifications')
