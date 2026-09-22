#!/usr/bin/env python3
"""Preserve all qualification attempts; before captures precede source mutation."""
import hashlib,json,subprocess,sys
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def read(p):return json.loads(p.read_text())
def write(p,x):p.write_text(json.dumps(x,indent=2)+'\n')
def static(x):return {k:v for k,v in x.items() if k not in ['samples','warmups','samples_requested','timing_claim','allocator_instrumented']}
def run(variant,lane,case,path,label,samples,warmups):
 b=read(P/f'{variant}-build.json')['binaries'][lane];assert sha(Path(b['path']))==b['sha256']
 out=P/'qualification-raw';out.mkdir(exist_ok=True);dest=out/(label+'.json');assert not dest.exists()
 cmd=['taskset','-c','12',b['path'],'--case',case,'--input',path,'--operation','format','--samples',str(samples),'--warmups',str(warmups)]
 with dest.open('wb') as o,(out/(label+'.stderr')).open('wb') as e:r=subprocess.run(cmd,cwd=ROOT,stdout=o,stderr=e)
 receipt=dict(command=cmd,exit_code=r.returncode,output=str(dest.relative_to(P)),sha256=sha(dest),stderr_sha256=sha(out/(label+'.stderr')))
 write(out/(label+'.receipt.json'),receipt);print(label,r.returncode,flush=True)
 if r.returncode:return None
 x=read(dest);assert all(s['oracle']==x['expected_oracle'] and s['output_inventory']==x['expected_output_inventory'] and s['output_sha256']==x['expected_output_sha256'] for s in x['samples'])
 return x,str(dest.relative_to(P))
def baseline():
 b=read(P/'baseline-build.json');assert all(sha(ROOT/f)==h for f,h in b['source'].items())
 x,out=run('baseline','native','ppt45543','test-data/poi/test-data/slideshow/45543.ppt','baseline-primary-native',50,3)
 assert static(x)==static(read(P.parent/'change-0731/oracle.json')['expected'])
 cases=[dict(id='primary',case='ppt45543',path=x['input'],bytes=(ROOT/x['input']).stat().st_size,sha256=sha(ROOT/x['input']))];oracles={'primary':dict(expected=x,qualification_output=out,sha256=sha(P/out))}
 candidates=['test-data/poi/test-data/slideshow/41246-1.ppt','test-data/office-interop/libreoffice-resaved/45543-transition-litchi.ppt','test-data/ole/ppt/SampleShow.ppt']
 for i,path in enumerate(candidates):
  result=run('baseline','native','ppt-secondary',path,f'baseline-secondary-{i}-native',50,3)
  if result:
   x,out=result;cases.append(dict(id='secondary',case='ppt-secondary',path=path,bytes=(ROOT/path).stat().st_size,sha256=sha(ROOT/path)));oracles['secondary']=dict(expected=x,qualification_output=out,sha256=sha(P/out));break
 assert len(cases)==2
 for c in cases:
  x,_=run('baseline','allocation',c['case'],c['path'],'baseline-'+c['id']+'-allocation',1,0);assert static(x)==static(oracles[c['id']]['expected'])
 assert all(sha(ROOT/f)==h for f,h in b['source'].items())
 write(P/'cases.json',cases);write(P/'oracle.json',oracles)
 write(P/'before.json',dict(status='passed',source=b['source'],build_sha256=sha(P/'baseline-build.json'),files={str(p.relative_to(P)):sha(p) for p in sorted((P/'qualification-raw').iterdir())},candidate_order=candidates))
def candidate():
 for c in read(P/'cases.json'):
  for lane in ['native','allocation']:
   x,_=run('candidate',lane,c['case'],c['path'],'candidate-'+c['id']+'-'+lane,1,0);assert static(x)==static(read(P/'oracle.json')[c['id']]['expected'])
 write(P/'qualification.json',dict(status='passed',files={str(p.relative_to(P)):sha(p) for p in sorted((P/'qualification-raw').iterdir())},builds={v:sha(P/f'{v}-build.json') for v in ['baseline','candidate']}))
if __name__=='__main__': {'baseline':baseline,'candidate':candidate}[sys.argv[1]]()
