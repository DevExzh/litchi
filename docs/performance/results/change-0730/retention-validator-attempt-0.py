#!/usr/bin/env python3
"""Build/capture/check the separate candidate-only retention observation lane."""
import hashlib,json,os,shutil,subprocess,sys,time
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3];TARGET=ROOT.parent/'litchi-target-0730';BIN=ROOT.parent/'litchi-0730-bin'
def read(p):return json.loads(p.read_text())
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def write(p,x):p.write_text(json.dumps(x,indent=2)+'\n')
def probes():return {str(p.relative_to(P)):sha(p) for p in (P/'retention-probe').rglob('*') if p.is_file()}
mode=sys.argv[1];assert mode in ['build','run','check']
if mode=='build':
 expected_source=read(P/'candidate-builds.json')['source_sha256'];assert all(sha(ROOT/rel)==h for rel,h in expected_source.items())
 i=0
 while (P/f'retention-build-{i}.json').exists():i+=1
 shutil.copytree(P/'retention-probe',P/f'retention-build-{i}-probe')
 cmd=['cargo','build','--manifest-path',str(P/'retention-probe/Cargo.toml'),'--release','--offline','--bins']
 if (P/'retention-probe/Cargo.lock').exists():cmd.append('--locked')
 start=time.monotonic()
 with (P/f'retention-build-{i}.log').open('w') as f:r=subprocess.run(cmd,cwd=ROOT,env=dict(os.environ,CARGO_TARGET_DIR=str(TARGET),CARGO_BUILD_JOBS='2'),stdout=f,stderr=subprocess.STDOUT)
 assert all(sha(ROOT/rel)==h for rel,h in expected_source.items())
 rows=[]
 if not r.returncode:
  for name in ['doc_retention_probe','doc_retention_probe_alloc']:
   dest=BIN/name;shutil.copy2(TARGET/'release'/name,dest);rows.append(dict(path=str(dest),bytes=dest.stat().st_size,sha256=sha(dest)))
 receipt=dict(command=cmd,exit_code=r.returncode,seconds=time.monotonic()-start,probe_sha256=probes(),binaries=rows,candidate_builds_sha256=sha(P/'candidate-builds.json'))
 write(P/f'retention-build-{i}.json',receipt);assert r.returncode==0;write(P/'retention-builds.json',receipt);print('PASS retention build');sys.exit()
b=read(P/'retention-builds.json');assert b['probe_sha256']==probes() and b['candidate_builds_sha256']==sha(P/'candidate-builds.json')
cases=read(P/'cases.json');out=P/'retention-captures';schedule=[(lane,c,edits,retention) for lane in ['ordinary','allocation'] for c in cases for edits in ['one','two'] for retention in ['zero','default','release']]
if mode=='run':
 out.mkdir(exist_ok=False);m=dict(status='running',builds_sha256=sha(P/'retention-builds.json'),runs=[])
 for lane,c,edits,retention in schedule:
  row=next(r for r in b['binaries'] if Path(r['path']).name=='doc_retention_probe'+('_alloc' if lane=='allocation' else ''));assert sha(Path(row['path']))==row['sha256']
  name=f'{lane}-{c["case"]}-{edits}-{retention}.json';cmd=['taskset','-c','12',row['path'],'--case',c['case'],'--input',c['path'],'--edits',edits,'--retention',retention]
  with (out/name).open('wb') as f,(out/(name+'.stderr')).open('wb') as e:r=subprocess.run(cmd,cwd=ROOT,stdout=f,stderr=e)
  m['runs'].append(dict(lane=lane,case=c['case'],edits=edits,retention=retention,command=cmd,exit_code=r.returncode,output=name,sha256=sha(out/name),stderr_sha256=sha(out/(name+'.stderr'))));write(out/'manifest.json',m);assert r.returncode==0
 m['status']='complete';write(out/'manifest.json',m)
m=read(out/'manifest.json');assert m['status']=='complete' and m['builds_sha256']==sha(P/'retention-builds.json') and len(m['runs'])==24
identities={};alloc={};rows=[]
for r,(lane,c,edits,retention) in zip(m['runs'],schedule,strict=True):
 assert (r['lane'],r['case'],r['edits'],r['retention'],r['exit_code'])==(lane,c['case'],edits,retention,0)
 assert sha(out/r['output'])==r['sha256'] and sha(out/(r['output']+'.stderr'))==r['stderr_sha256'];x=read(out/r['output']);assert x['source_sha256']==c['sha256'] and x['source_bytes']==c['bytes']
 assert x['direct_output_equal_reference'] and x['output_sha256']==x['reference_output_sha256'];key=(c['case'],edits);assert x['output_sha256']==identities.setdefault(key,x['output_sha256'])
 assert x['retention']==retention and x['route']==edits and x['allocator_instrumented']==(lane=='allocation');count=1 if edits=='one' else 2;assert x['edits']==count
 assert x['retention_ceiling_bytes']==(0 if retention=='zero' else 8*1024*1024)
 for i,s in enumerate(x['stages'][:count]):
  assert s['stage']==i+1
  if i==0 or retention!='default':assert s['before'] is None
  else:assert s['before']==x['stages'][i-1]['after']
  if retention=='zero':assert s['after'] is None
  else:assert 0<s['after']<=x['retention_ceiling_bytes']
  assert s['after_release'] is None
 assert x['final_retained_before_commit']==(x['stages'][count-1]['after'] if retention=='default' else None)
 if lane=='allocation':alloc[(c['case'],edits,retention)]=x['allocation']
 else:assert x['allocation'] is None
 rows.append(dict(lane=lane,case=c['case'],edits=edits,retention=retention,stages=x['stages'],final_retained_before_commit=x['final_retained_before_commit'],allocation=x['allocation']))
comparisons=[]
for c in cases:
 for edits in ['one','two']:
  base=alloc[(c['case'],edits,'zero')]
  for retention in ['default','release']:
   value=alloc[(c['case'],edits,retention)];changes={k:100*(value[k]/base[k]-1) for k in base if base[k]};comparisons.append(dict(case=c['case'],edits=edits,retention=retention,percent_changes=changes,peak_live_flag_over_5pct=changes['peak_live_bytes']>5))
result=dict(processes=rows,comparisons=comparisons)
if mode=='run':write(P/'retention-analysis.json',result)
else:assert result==read(P/'retention-analysis.json')
print('PASS 24 retention processes; direct output equality and bounded capacity; no RSS/timing claim')
