"""Replay supplemental count/layout observations, without latency inference."""
import argparse
import custody as c
from profile_analysis import artifact

COUNTS=[0,1,2,3,4,5,8,9,16,17,31,32,33,64]

def analyze():
 reports={}; normalized={}
 for leg in ['before','after']:
  p=c.P/f'attribute-{leg}'
  frozen=c.read(p/'frozen-inputs.json')
  for name,sha in frozen.items():assert c.sha(c.P/name)==sha
  build=c.read(p/'build.json');row=c.read(p/'receipt.json');cleanup=c.read(p/'cleanup.json')
  assert build['exit_code']==row['exit_code']==0
  artifact(build['log']);artifact(row['log']);artifact(row['lock'])
  assert c.read(artifact(row['source']))==c.read(c.P/f'build-{leg}/source.json')
  assert cleanup=={'removed_binary':row['binary'],'binary_removed':True}
  from pathlib import Path
  assert not Path(row['binary']['path']).exists()
  expected=['cargo','build','--release','--offline','--manifest-path',str(c.P/'attribute-probe-src/Cargo.toml')]
  if leg=='after':expected+=['--locked']
  origin=c.read(c.P/'origin.json')['main']
  assert [x.replace(origin,str(c.ROOT)) for x in build['command']]==expected
  assert row['command']==['taskset','-c','12',row['binary']['path'],row['report']['path']]
  assert build['started']<=build['ended']<=row['started']<=row['ended']
  report=c.read(artifact(row['report']));reports[leg]=report
  assert report['schema']=='litchi.performance.0794.attribute-diagnostic.v1'
  assert 'no timing measurement' in report['scope']
  assert [(r['attributes'],r['repeat']) for r in report['rows']]==[(n,i) for n in COUNTS for i in range(3)]
  values={}
  for r in report['rows']:
   assert r['seen']==r['attributes'];a=r['allocation']
   assert a['status']=='measured' and a['failed_allocation_calls']==0
   v={**{k:a[k] for k in ['allocation_calls','reallocation_calls','deallocation_calls','allocated_bytes','deallocated_bytes']},
      'net_live_bytes':a['live_bytes_after']-a['live_bytes_before'],
      'peak_above_entry_bytes':a['region_peak_live_bytes']-a['live_bytes_before']}
   assert v['net_live_bytes']==0
   values.setdefault(r['attributes'],v);assert values[r['attributes']]==v
  normalized[leg]=values
 assert c.read(c.P/'attribute-before/frozen-inputs.json')==c.read(c.P/'attribute-after/frozen-inputs.json')
 assert c.read(c.P/'attribute-before/receipt.json')['lock']==c.read(c.P/'attribute-after/receipt.json')['lock']
 rows=[]
 for n in COUNTS:
  before,after=normalized['before'][n],normalized['after'][n]
  rows.append({'attributes':n,'before':before,'after':after,
               'increases':[k for k in ['allocation_calls','allocated_bytes','net_live_bytes','peak_above_entry_bytes'] if after[k]>before[k]]})
 return {'schema':'litchi.performance.0794.attribute-analysis.v1','reports':2,'iterations':84,'rows':rows,
         'iterator_size_bytes':{leg:r['iterator_size_bytes'] for leg,r in reports.items()},
         'any_resource_increase':any(r['increases'] for r in rows),'latency_claim':False,
         'scope':'Supplemental helper counts and layout only, outside public-workflow adoption policy'}

if __name__=='__main__':
 parser=argparse.ArgumentParser();parser.add_argument('--write',action='store_true');parser.add_argument('--check',action='store_true')
 args=parser.parse_args();result=analyze()
 if args.write:c.write(c.P/'attribute-analysis.json',result)
 if args.check:assert c.read(c.P/'attribute-analysis.json')==result
 print('0794 helper replay PASS; resources increased:',result['any_resource_increase'],'sizes:',result['iterator_size_bytes'])
