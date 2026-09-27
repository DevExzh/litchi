"""Replay ordinary native controls separately from guest instruction profiles."""
import math,random,statistics,sys
from pathlib import Path
import custody as c

def quantile(values,q):return sorted(values)[max(0,math.ceil(len(values)*q)-1)]
def interval(values,plan):
 rng=random.Random(plan['bootstrap']['seed']);n=plan['bootstrap']['resamples']
 v=sorted(statistics.median(rng.choices(values,k=len(values))) for _ in range(n))
 return [v[math.floor(.025*n)],v[math.ceil(.975*n)-1]]
def relocate(path):
 marker='/change-0784/'
 return c.P/path.split(marker,1)[1] if marker in path else Path(path)
def artifact(r):
 p=relocate(r['path']);assert p.is_file() and p.stat().st_size==r['bytes'] and c.sha(p)==r['sha256'],p
 return p
def binaries(build):
 cleanup=c.read(c.P/'cleanup.json') if (c.P/'cleanup.json').exists() else None
 for leg,r in build['binaries'].items():
  if Path(r['path']).exists():artifact(r)
  else:assert cleanup and cleanup['verified_before_removal'] is True and cleanup['binaries'][leg]==r

def oracle(shape):
 old=c.P.parent/'change-0780';name=f'qualification/0-{shape}-capture-before.json'
 assert c.sha(old/name)==c.read(old/'seal.json')['files'][name]
 d=c.read(old/name);return d['source'],d['samples'][0]['output'],d['samples'][0]['verification']

def report(path,shape,leg,n,warmup,build):
 d=c.read(path);source,out,verification=oracle(shape)
 assert d['schema']=='litchi.pptx.capture-profile-probe.v1' and d['tool']=='pptx-capture-probe-0784'
 assert d['mode']=='capture' and d['shape']==shape and d['timing_scope']=='Package::opened_presentation only'
 dims={'tiny':(3,4),'medium':(12,8),'large':(100,100)}[shape]
 assert (d['slides'],d['shapes_per_slide'])==dims
 assert d['marker']=='litchi-perf-0780-static-mce-capabilities'
 assert d['source']==source and d['warmup']==warmup and d['samples_requested']==n and len(d['samples'])==n
 assert d['allocator']=={'binary':Path(build['binaries'][leg]['path']).name,'allocator':'Rust system allocator','instrumentation':'none','counter_revision':None}
 values=[]
 for i,s in enumerate(d['samples']):
  assert s['index']==i and type(s['elapsed_ns']) is int and s['elapsed_ns']>0
  assert s['output']==out and s['verification']==verification and s['source_sha256']==source['sha256']
  assert s['metrics']=={'elapsed_ns':s['elapsed_ns'],'slides':dims[0],'shapes_per_slide':dims[1],'captured_slides':dims[0],'captured_shapes_per_slide':dims[1]}
  assert 'allocation' not in s or s['allocation'] is None
  values.append(s['elapsed_ns'])
 return values

def analyze():
 plan=c.read(c.P/'plan.json');build=c.read(c.P/'build/build.json');binaries(build)
 complete=c.read(c.P/'native/complete.json');assert complete=={'processes':36,'plan_sha256':c.sha(c.P/'plan.json'),'build_sha256':c.sha(c.P/'build/build.json')}
 rows=c.read(c.P/'native/receipts.json');jobs=[(b,s,l) for b,order in enumerate(plan['native']['orders']) for s in plan['native']['shapes'] for l in order]
 assert len(rows)==len(jobs)==36
 processes=[];lookup={};previous=0
 for row,(block,shape,leg) in zip(rows,jobs):
  assert (row['block'],row['shape'],row['leg'])==(block,shape,leg) and row['exit_code']==0
  assert previous<=row['started']<=row['ended'];previous=row['ended']
  rp=artifact(row['report']);rss=artifact(row['rss']);artifact(row['log'])
  expected=['/usr/bin/time','-f','%M','-o',row['rss']['path'],'taskset','-c',str(plan['cpu']),build['binaries'][leg]['path'],'--mode','capture','--shape',shape,'--samples','30','--warmup','3','--output',row['report']['path']]
  assert row['command']==expected
  values=report(rp,shape,leg,30,3,build);mem=int(rss.read_text());assert mem>0
  item={'block':block,'shape':shape,'leg':leg,'p50_ns':quantile(values,.5),'p95_ns':quantile(values,.95),'p99_ns':quantile(values,.99),'peak_process_rss_kib':mem,'report':rp.name}
  processes.append(item);lookup[(block,shape,leg)]=item
 summaries={};flags=[]
 for shape in plan['native']['shapes']:
  groups={}
  for leg in ['control','profile']:
   group={}
   for key in ['p50_ns','p95_ns','p99_ns','peak_process_rss_kib']:
    v=[lookup[(b,shape,leg)][key] for b in range(6)];spread=(max(v)/min(v)-1)*100
    group[key]={'median':statistics.median(v),'values':v,'spread_percent':spread}
    if spread>5:flags.append({'shape':shape,'leg':leg,'metric':key,'spread_percent':spread})
   groups[leg]=group
  ratios=[lookup[(b,shape,'profile')]['p50_ns']/lookup[(b,shape,'control')]['p50_ns'] for b in range(6)]
  groups['wrapper_control']={'ratios':ratios,'median_ratio':statistics.median(ratios),'pairs_over_5_percent':[i for i,r in enumerate(ratios) if r>1.05],'bootstrap95':interval(ratios,plan),'interpretation':'Wrapper/codegen perturbation; not a production speedup or Callgrind cost'}
  summaries[shape]=groups
 return {'processes':processes,'summaries':summaries,'spread_flags':flags,'verified_measured_samples':1080,'scope':'Native capture-only latency; RSS whole process; all profiles separate'}

def main():
 assert sys.argv[1:] in [['--write'],['--check']]
 result=analyze();encoded=__import__('json').dumps(result,indent=2,sort_keys=True)+'\n';path=c.P/'native-analysis.json'
 if sys.argv[1]=='--write':path.write_text(encoded)
 else:assert path.read_text()==encoded
 print('Native capture controls replay PASS')
if __name__=='__main__':main()
