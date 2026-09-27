"""Offline heap-allocation reconciliation, separate from native adoption policy."""
import argparse
from collections import Counter
from pathlib import Path
import re
import custody as c

P=c.P
OWNER=re.compile(r'^namespace_uri_probe::capture_region_0793::h[0-9a-f]+(?: \(.*\))?$')

def artifact(a):
 path=Path(a['path'])
 if path.is_absolute():path=c.ROOT/path.relative_to(c.read(P/'origin.json')['main'])
 else:path=P/path
 assert c.sha(path)==a['sha256'] and path.stat().st_size==a['bytes'],path
 return path

def stacks(path):
 out=Counter()
 for line in path.read_text().splitlines():
  stack,n=line.rsplit(' ',1);assert n.isdecimal() and int(n)>0
  out[stack]+=int(n)
 return out

def analyze():
 plan=c.read(P/'plan.json');done=c.read(P/'profile/complete.json')
 assert done['children']==4 and done['decodes']==8
 for field in ['receipts','decode_receipts','source']:artifact(done[field])
 assert c.read(artifact(done['source']))==c.read(P/'build-after/source.json')
 rows=c.read(artifact(done['receipts']));decodes=c.read(artifact(done['decode_receipts']))
 assert [(r['repeat'],r['leg']) for r in rows]==[(i,leg) for i,order in enumerate(plan['profile']['orders']) for leg in order]
 assert len(decodes)==8
 expected={}
 for leg in ['before','after']:
  values=[s['allocation']['allocation_calls'] for i in range(2) for s in c.read(P/f'allocation/{i}-large-capture-{leg}.json')['samples']]
  assert len(set(values))==1;expected[leg]=values[0]
 output=[]
 for i,row in enumerate(rows):
  leg=row['leg'];repeat=row['repeat'];build=c.read(P/f'build-{leg}/build.json')
  assert row['exit_code']==0 and row['binary']==build['binaries']['profile']
  assert row['ended']>=row['started'];artifact(row['log'])
  assert len(row['traces'])==1;trace=row['traces'][0];artifact(trace)
  raw=c.read(artifact(row['report']));ref=c.read(P/f'qualification/0-large-capture-before.json')
  for key in ['schema','tool','mode','shape','slides','shapes_per_slide','timing_scope','marker','source','fixture']:
   assert raw[key]==ref[key],key
  assert raw['samples_requested']==len(raw['samples'])==1 and raw['warmup']==0
  sample=raw['samples'][0]
  assert 'allocation' not in sample
  for key in ['source_sha256','output','verification']:assert sample[key]==ref['samples'][0][key]
  cmd=['taskset','-c',str(plan['cpu']),'heaptrack','--record-only','-o',row['report']['path'].removesuffix('.json')+'.heaptrack',row['binary']['path'],'--mode','capture','--shape','large','--samples','1','--warmup','0','--output',row['report']['path']]
  assert row['command']==cmd
  scopes={}
  for d,scope in zip(decodes[2*i:2*i+2],['whole','owner']):
   assert (d['repeat'],d['leg'],d['scope'])==(repeat,leg,scope)
   assert d['trace']==trace and d['exit_code']==0
   assert artifact(d['stderr']).stat().st_size==0
   cmd=['heaptrack_print','-f',trace['path'],'-m','0','-t','0','-p','1','-a','1','-T','0','-n','20','--flamegraph-cost-type','allocations','-F',d['stacks']['path'],'-H',d['histogram']['path']]
   if scope=='owner':cmd+=['--filter-bt-function','capture_region_0793']
   assert d['command']==cmd
   stack=stacks(artifact(d['stacks']))
   hist=[tuple(map(int,l.split())) for l in artifact(d['histogram']).read_text().splitlines()]
   assert all(len(v)==2 and v[0]>=0 and v[1]>0 for v in hist)
   summary=int(re.search(r'calls to allocation functions:\s*(\d+)',artifact(d['log']).read_text())[1])
   scopes[scope]=(stack,sum(n for _,n in hist),summary)
  whole,owner=scopes['whole'][0],scopes['owner'][0]
  total=sum(whole.values());observed=sum(owner.values())
  selected={s:n for s,n in whole.items() if any(OWNER.fullmatch(f) for f in s.split(';'))}
  checks={'whole_conservation':total==scopes['whole'][1]==scopes['whole'][2],
          'owner_subset':owner==selected,'owner_public_descendant':observed>0 and all('opened_presentation' in s for s in owner),
          'counter_match':observed==expected[leg]}
  output.append({'repeat':repeat,'leg':leg,'whole_calls':total,'owner_calls':observed,'expected_operation_calls':expected[leg],
                 'checks':checks,'qualified':all(checks.values()),
                 'filtered_histogram_is_global':scopes['owner'][1]==total,'filtered_summary_is_global':scopes['owner'][2]==total,
                 'nested_quick_xml_duplicate_check':sum(n for s,n in owner.items() if 'check_for_duplicates' in s),
                 'nested_notes_inspector':sum(n for s,n in owner.items() if 'notes::codec::inspect_element' in s),
                 'nested_checked_attributes':sum(n for s,n in owner.items() if 'xml_attributes::' in s)})
 assert all(a['ended']<=b['started'] for a,b in zip(rows,rows[1:]))
 assert rows[-1]['ended']<=decodes[0]['started']
 assert all(a['ended']<=b['started'] for a,b in zip(decodes,decodes[1:]))
 return {'schema':'litchi.performance.0794.profile-analysis.v1','reports':4,'samples':4,'decodes':8,
         'rows':output,'qualified':all(r['qualified'] for r in output),'timing_claim':False,'rss_claim':False,
         'counter_semantics':'allocation_calls includes successful reallocations; do not add reallocation_calls',
         'nested_costs_overlap':True}

if __name__=='__main__':
 parser=argparse.ArgumentParser();parser.add_argument('--write',action='store_true');parser.add_argument('--check',action='store_true')
 args=parser.parse_args();result=analyze()
 if args.write:c.write(P/'profile-analysis.json',result)
 if args.check:assert c.read(P/'profile-analysis.json')==result
 print('0794 profile replay PASS; qualification',result['qualified'])
