"""Replay raw paired measurements without treating changed outcomes as equivalent."""
from pathlib import Path
import hashlib,json,statistics,math
P=Path(__file__).resolve().parent
D=P/'measure-0'
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def read(p):return json.loads(p.read_text())
def summary(xs):
 m=statistics.median(xs)
 return {'median':m,'min':min(xs),'max':max(xs),'spread_percent':100*(max(xs)-min(xs))/m}
def analyze():
 complete=read(D/'complete.json');runs=read(D/'runs.json');assert complete['runs']==len(runs)
 groups={};differentials={}
 for r in runs:
  assert r['exit']==0
  for field in ['report','log','rss']:
   if field in r:assert sha(D/r[field])==r[field+'_sha256']
  report=read(D/r['report'])
  if r['kind']=='differential':differentials[r['leg']]=report;continue
  if r['kind']!='adversarial':continue
  assert len(report['elapsed_ns'])==report['samples']
  assert report['p50_ns']==sorted(report['elapsed_ns'])[len(report['elapsed_ns'])//2]
  assert len(report['outcomes'])==1
  assert report['outcomes'][0].startswith('ok:'),(r['case'],report['outcomes'])
  groups.setdefault(r['case'],{}).setdefault(r['leg'],[]).append((report,int((D/r['rss']).read_text())))
 cases={}
 for case,legs in groups.items():
  assert set(legs)=={'before','after'}
  identity={(x['input_bytes'],x['input_sha256'],tuple(x['outcomes']),x['samples'],x['warmup']) for leg in legs.values() for x,_ in leg}
  assert len(identity)==1,(case,'changed input or outcome')
  item={}
  for leg,rows in legs.items():
   assert len(rows)==3
   item[leg]={'p50_ns':summary([x['p50_ns'] for x,_ in rows]),'rss_kib':summary([rss for _,rss in rows]),'sample_max_ns':max(max(x['elapsed_ns']) for x,_ in rows)}
  for metric in ['p50_ns','rss_kib']:
   ratio=item['after'][metric]['median']/item['before'][metric]['median']
   item[metric+'_change_percent']=100*(ratio-1)
   item[metric+'_regression_flag']=ratio>1.05
  item['input_bytes'],item['input_sha256'],outcomes,item['samples'],item['warmup']=next(iter(identity))
  item['outcome']=outcomes[0];cases[case]=item
 before,after=differentials['before'],differentials['after']
 assert {k:v for k,v in before.items() if k!='results'}=={k:v for k,v in after.items() if k!='results'}
 assert before['results'].keys()==after['results'].keys()
 changes={k:{'before':v,'after':after['results'][k]} for k,v in before['results'].items() if v!=after['results'][k]}
 result={'cases':cases,'differential':{'summary':{k:v for k,v in before.items() if k!='results'},'comparisons':len(before['results']),'unchanged':len(before['results'])-len(changes),'changes':changes},'limits':'Three processes/leg. Small sample maxima are not stable p95/p99 estimates. RSS includes process/input/observer overhead. Changes require explicit triage; this script does not accept them.'}
 return result
if __name__=='__main__':
 result=analyze()
 (P/'analysis.json').write_text(json.dumps(result,indent=2)+'\n')
 print(json.dumps({'cases':len(result['cases']),'differential_changes':len(result['differential']['changes'])}))
