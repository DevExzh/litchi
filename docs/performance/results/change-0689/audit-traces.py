#!/usr/bin/env python3
"""Audit diagnostic-only chain route traces separately from latency evidence."""
import hashlib,json,re,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
def read(p):return json.loads(p.read_text())
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def routes(path):
 result=[];current=None;pending=None
 for line in path.read_text().splitlines():
  if line.startswith('TRACE replay '):
   current={'target':line,'worksheet':[],'sst':[],'locations':[]};result.append(current);pending=None
  elif current is not None:
   if line.startswith('TRACE worksheet '):pending='worksheet';current['locations'].append(line)
   elif line.startswith('TRACE sst '):pending='sst';current['locations'].append(line)
   elif line.startswith('TRACE chain '):
    assert pending is not None
    m=re.fullmatch(r'TRACE chain table=(\w+) from=(\d+) to=(\d+) links=(\d+)',line);assert m,line
    table,a,b,n=m.groups();a,b,n=map(int,[a,b,n]);assert n==max(b-a,0)
    current[pending].append(dict(table=table,start=a,end=b,links=n))
 return result
for phase,folder in [('baseline','route-trace'),('candidate','candidate-route-trace')]:
 d=P/folder;m=read(d/'manifest.json')
 for n,h in m['raw_sha256'].items():assert sha(d/n)==h,n
 for c in m['commands']:assert c['exit_code']==c.get('expected_exit_code',0)
 builds=read(P/(phase+'-builds.json'));frozen=builds[0]['source_sha256']
 assert m['binary_sha256']!=builds[0]['binary_sha256']
 changed={'crates/litchi-cfb/src/shared.rs','crates/litchi-xls/src/workbook/source.rs'}
 if phase=='candidate':
  assert m['original_source_sha256']==frozen
  assert {n for n,h in m['instrumented_source_sha256'].items() if h!=frozen[n]}==changed
 else:
  assert m['baseline_head']==read(P/'baseline.json')['baseline_head']
  assert set(m['instrumented_source_sha256'])==changed
  assert all(h!=frozen[n] for n,h in m['instrumented_source_sha256'].items())
rows=[]
for c in read(P/'cases.json'):
 if c['budget']!=2097152:continue
 name=c['case'];b=routes(P/'route-trace'/(name+'.trace'));a=routes(P/'candidate-route-trace'/(name+'.trace'))
 if name.startswith('formula-refusal'):
  assert not a and not b
  for folder in ['route-trace','candidate-route-trace']:assert 'shared Formula metadata requires a leading PtgExp token' in (P/folder/(name+'.trace')).read_text()
 else:
  assert len(a)==len(b)==2
  for old,new in zip(b,a):
   assert old['target']==new['target'] and old['locations']==new['locations']
   assert old['worksheet']==new['worksheet']
   assert len(old['sst'])==len(new['sst'])
   for x,y in zip(old['sst'],new['sst']):assert x['table']==y['table'] and x['end']==y['end']==y['start'] and y['links']==0
 rows.append(dict(case=name,baseline=b,candidate=a))
(P/'route-comparison.json').write_text(json.dumps(rows,indent=2)+'\n')
print('PASS diagnostic routes: same worksheet walks/locations; selected repeated SST walks reach zero links. Formula refusal remains before replay. Traced timings excluded.')
