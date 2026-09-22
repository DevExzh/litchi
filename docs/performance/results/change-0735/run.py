#!/usr/bin/env python3
"""Frozen serial, paired ordinary public-workflow capture."""
import hashlib,json,subprocess,sys,time
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
def read(p):return json.loads(p.read_text())
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def write(p,x):p.write_text(json.dumps(x,indent=2)+'\n')
def guard():
 builds={v:read(P/f'{v}-build.json') for v in ['baseline','candidate']};b=builds['candidate']
 assert {str(f.relative_to(ROOT)):sha(f) for f in sorted(list((ROOT/'crates').rglob('*.rs'))+list((ROOT/'crates').rglob('Cargo.toml'))+[ROOT/'Cargo.toml',ROOT/'Cargo.lock'])}==b['source']
 assert all(sha(P/f)==h for f,h in b['probe'].items())
 assert builds['baseline']['probe']==b['probe']
 assert [f for f in b['source'] if b['source'][f]!=builds['baseline']['source'][f]]==[read(P/'base.json')['owned_file']]
 assert all(sha(ROOT/f)==h for f,h in read(P/'constraints.json').items())
 q=read(P/'quality.json');assert q['source']==b['source'] and q['probe']==b['probe'];assert len(q['runs'])==7
 for v,x in builds.items():
  assert sha(P/x['quality'])==x['quality_sha256'];z=read(P/x['quality']);assert z['source']==x['source'] and z['probe']==x['probe'] and len(z['runs'])==5
  for r in z['runs']:assert r['exit_code']==0 and sha(P/r['output'])==r['sha256']
  for r in x['binaries'].values():assert sha(Path(r['path']))==r['sha256']
 for r in q['runs']:assert r['exit_code']==0 and sha(P/r['output'])==r['sha256']
 return builds

def bindings():
 builds=guard();paths=[]
 for name in ['run.py','build.py','quality.py','qualify.py','analyze.py','audit.py','preflight.py','plan.json','hypothesis.md','clone-hypothesis.json','source-equivalence.json','source-guard.py','environment.json','constraints.json','base.json','cases.json','oracle.json','before.json','baseline-build.json','candidate-build.json','quality.json','qualification.json','source-review.json','source-review.md']:
  paths.append(P/name)
 for folder in ['probe','qualification-raw','source-archive','candidate']:
  paths+=list((P/folder).rglob('*'))
 for folder in list(P.glob('build-*'))+list(P.glob('quality-*')):
  if folder.is_dir():paths+=list(folder.rglob('*'))
 paths += [ROOT/f for f in builds['candidate']['source']]
 paths += [ROOT/c['path'] for c in read(P/'cases.json')]
 for b in builds.values():paths += [Path(r['path']) for r in b['binaries'].values()]
 return {str(p):sha(p) for p in paths if p.is_file()}
def command(spec):
 c=next(c for c in read(P/'cases.json') if c['id']==spec['case']);b=read(P/f'{spec["variant"]}-build.json')['binaries'][spec['lane']]
 return ['taskset','-c','12',b['path'],'--case',c['case'],'--input',c['path'],'--operation','format','--samples','50' if spec['lane']=='native' else '1','--warmups','3' if spec['lane']=='native' else '0']
def main():
 mode=sys.argv[1];assert mode in ['freeze','capture'];guard()
 if mode=='freeze':
  assert read(P/'qualification.json')['status']=='passed';assert not (P/'freeze.json').exists();write(P/'freeze.json',bindings());print('frozen');return
 assert read(P/'freeze.json')==bindings()
 preflight=read(P/'preflight.json');assert preflight['status']=='passed' and preflight['freeze_sha256']==sha(P/'freeze.json')
 for name in ['preflight.py','analyze.py','audit.py']:assert preflight['scripts'][name]==sha(P/name)
 out=P/'captures';out.mkdir();rows=[]
 for spec in read(P/'plan.json')['schedule']:
  name=f'{spec["lane"]}-c{spec["cycle"]}-r{spec["repeat"]}-{spec["case"]}-{spec["variant"]}.json';dest=out/name;cmd=command(spec);start=time.monotonic()
  with dest.open('wb') as o,Path(str(dest)+'.stderr').open('wb') as e:r=subprocess.run(cmd,cwd=ROOT,stdout=o,stderr=e)
  row=dict(spec,command=cmd,exit_code=r.returncode,seconds=time.monotonic()-start,output=name,sha256=sha(dest),stderr_sha256=sha(Path(str(dest)+'.stderr')));rows.append(row)
  write(out/'manifest.json',dict(status='running',freeze_sha256=sha(P/'freeze.json'),preflight_sha256=sha(P/'preflight.json'),runs=rows));print(name,r.returncode,flush=True);assert r.returncode==0
 assert read(P/'freeze.json')==bindings();guard();write(out/'manifest.json',dict(status='complete',freeze_sha256=sha(P/'freeze.json'),preflight_sha256=sha(P/'preflight.json'),runs=rows))
if __name__=='__main__':main()
