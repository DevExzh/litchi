#!/usr/bin/env python3
"""Offline exact-oracle, native statistics and Callgrind edge replay."""
import hashlib,json,math,re,statistics
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
def read(p):return json.loads(p.read_text())
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def stats(v):
 s=sorted(v);return dict(n=len(s),p50=statistics.median(s),mean=statistics.fmean(s),p95=s[math.ceil(.95*len(s))-1],p99=s[math.ceil(.99*len(s))-1],maximum=s[-1])
def profile(path):
 ids={};self_cost={};edges={};current=None;callee=None;pending=False;summary=None;total=None
 for line in path.read_text().splitlines():
  if line.startswith('events:'):assert line=='events: Ir'
  if line.startswith('positions:'):assert line=='positions: line'
  if line.startswith('summary:'):summary=int(line.split()[1])
  if line.startswith('totals:'):total=int(line.split()[1])
  if line.startswith(('fn=','cfn=')):
   key,value=line.split('=',1);m=re.fullmatch(r'\((\d+)\)(?: (.*))?',value)
   if m:
    ident,name=m.groups()
    if name is not None:ids[ident]=name
    value=ids[ident]
   if key=='fn':current=value;pending=False
   else:callee=value
  elif line.startswith('calls='):pending=True
  elif line and line[0] in '*+-0123456789':
   fields=line.split();assert len(fields)==2,line;cost=int(fields[-1]);assert current is not None
   if pending:
    assert callee is not None;key=(current,callee);edges[key]=edges.get(key,0)+cost;pending=False
   else:self_cost[current]=self_cost.get(current,0)+cost
 assert summary is not None and summary>0
 assert sum(self_cost.values())==summary,(sum(self_cost.values()),summary)
 if total is not None:assert total==summary
 owner=read(P/'plan.json')['callgrind']['symbol'];assert owner in self_cost
 inclusive=self_cost[owner]+sum(v for (a,b),v in edges.items() if a==owner)
 assert inclusive==summary,(inclusive,summary)
 return dict(instructions=summary,owner=owner,owner_inclusive=inclusive,self_functions=sorted([dict(function=k,Ir=v,percent=100*v/summary) for k,v in self_cost.items() if v],key=lambda r:-r['Ir']),edges=sorted([dict(caller=a,callee=b,Ir=v,percent_of_owner=100*v/summary) for (a,b),v in edges.items() if v],key=lambda r:-r['Ir']))
def main():
 quality=read(P/'quality/manifest.json');assert len(quality)==5 and all(r['exit_code']==0 and sha(P/r['output'])==r['sha256'] for r in quality)
 build=read(P/'build.json');freeze=read(P/'freeze.json');cleanup=read(P/'cleanup.json') if (P/'cleanup.json').exists() else None
 for raw,h in freeze.items():assert sha(Path(raw))==h,raw
 for r in build['binaries']:
  p=Path(r['path'])
  if p.exists():assert sha(p)==r['sha256'] and p.stat().st_size==r['bytes']
  else:assert cleanup and cleanup['removed'] and r in cleanup['binaries']
 for rel,h in build['source'].items():assert sha(ROOT/rel)==h,rel
 for rel,h in read(P/'constraints.json').items():assert sha(ROOT/rel)==h,rel
 ancestry=read(P/'ancestry.json');assert sha(ROOT/ancestry['manifest_path'])==ancestry['manifest_sha256']
 contract=read(P/'oracle.json');assert read(ROOT/ancestry['manifest_path'])['files'][ancestry['oracle_entry']]['sha256']==contract['sha256'];assert sha(ROOT/contract['reference'])==contract['sha256'];expected=contract['expected'];m=read(P/'captures/manifest.json');assert m['status']=='complete' and m['freeze_sha256']==sha(P/'freeze.json') and len(m['runs'])==9
 rows=[];profiles=[];files={'manifest.json'}
 for i,r in enumerate(m['runs']):
  lane=['native','allocation','callgrind'][i%3];repeat=i//3;assert (r['lane'],r['repeat'],r['exit_code'])==(lane,repeat,0)
  path=P/r['output'];assert path==P/'captures'/f'{lane}-{repeat}.json';files.add(path.name);assert sha(path)==r['sha256'];x=read(path)
  for key in ['schema_version','case','format','operation','input','policy','policy_applied','source_sha256','expected_output_sha256','replacements_sha256','source_inventory','expected_output_inventory','replacements','changed_length_proof','expected_oracle','oracle_controls']:assert x[key]==expected[key],key
  count=50 if lane=='native' else 1;warmups=3 if lane=='native' else 0;assert len(x['samples'])==x['samples_requested']==count and x['warmups']==warmups
  assert x['allocator_instrumented']==(lane=='allocation') and x['timing_claim']==(lane!='allocation')
  for index,sample in enumerate(x['samples']):
   assert sample['index']==index
   if lane!='allocation':assert sample['phase_ns']['whole_ns']>0
   assert sample['output_sha256']==expected['expected_output_sha256']
   assert sample['output_inventory']==expected['expected_output_inventory']
   assert sample['oracle']==expected['expected_oracle']
  binary=build['binaries'][1 if lane=='allocation' else 0]['path'];cmd=['taskset','-c','12']
  if lane=='callgrind':cmd+=['valgrind','--tool=callgrind','--collect-atstart=no','--toggle-collect='+read(P/'plan.json')['callgrind']['symbol'],'--callgrind-out-file='+str(P/'captures'/f'{lane}-{repeat}.callgrind'),'--log-file='+str(P/'captures'/f'{lane}-{repeat}.valgrind')]
  cmd+=[binary,'--case','ppt45543','--input',read(P/'case.json')['path'],'--operation','format','--samples',str(count),'--warmups',str(warmups)];assert r['command']==cmd
  if lane=='native':rows.append(dict(repeat=repeat,lane=lane,timing_ns=stats([s['phase_ns']['whole_ns'] for s in x['samples']])))
  elif lane=='allocation':rows.append(dict(repeat=repeat,lane=lane,allocation=x['samples'][0]['allocations']))
  else:
   for suffix in ['callgrind','valgrind']:
    raw=P/'captures'/f'{lane}-{repeat}.{suffix}';assert sha(raw)==r[suffix+'_sha256'];files.add(raw.name)
   a=r['annotation'];assert a['exit_code']==0 and sha(P/a['output'])==a['sha256'];files.add(Path(a['output']).name)
   prof=profile(P/'captures'/f'{lane}-{repeat}.callgrind');annotation=(P/a['output']).read_text();tot=re.search(r'([\d,]+)\s+\([^\n]*\)\s+PROGRAM TOTALS',annotation);assert tot and int(tot[1].replace(',',''))==prof['instructions'];profiles.append(dict(repeat=repeat,**prof))
 assert {p.name for p in (P/'captures').iterdir()}==files
 native=[r for r in rows if r['lane']=='native'];controls=[]
 for r in native[1:]:
  d={k:100*(r['timing_ns'][k]/native[0]['timing_ns'][k]-1) for k in ['p50','mean','p95','p99','maximum']};controls.append(dict(repeat=r['repeat'],versus=0,percent=d,central_flags=[k for k in ['p50','mean'] if abs(d[k])>5]))
 result=dict(status='passed',processes=rows,native_controls=controls,profiles=profiles,scope='Callgrind Ir only; no instrumented latency/hardware/RSS claim')
 (P/'analysis.json').write_text(json.dumps(result,indent=2)+'\n');print('PASS nine processes, exact0728oracle, raw function/edge totals and source custody')
if __name__=='__main__':main()
