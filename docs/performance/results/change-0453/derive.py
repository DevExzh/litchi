#!/usr/bin/env python3
"""Derive normal latency and separate allocator evidence without pooling them."""
import gzip,json,random,re,statistics
from pathlib import Path
ROOT=Path(__file__).resolve().parent
def require(ok,message):
 if not ok:raise ValueError(message)
def raw(name):
 p=ROOT/name;return p.read_bytes() if p.exists() else gzip.decompress(p.with_name(p.name+'.gz').read_bytes())
def load(name):return json.loads(raw(name))
def quantile(v,q):
 v=sorted(v);x=(len(v)-1)*q;i=int(x);return v[i]+(v[min(i+1,len(v)-1)]-v[i])*(x-i)
def ci(v,seed):
 rng=random.Random(seed);b=[statistics.median(rng.choices(v,k=len(v))) for _ in range(2000)];return [quantile(b,.025),quantile(b,.975)]
def delta(a,b):return 0 if a==b==0 else None if a==0 else (b-a)/abs(a)*100
def derive():
 protocol=load('protocol.json');rows=[]
 for i,lane in enumerate(protocol['order']):
  report=load(f'runs/{i}/report.json');samples=report['samples_raw'];require(len(samples)==30,'30 samples');row={'lane':i,**lane}
  for phase in ['open','plan','publication','api_sum']:
   values=[x['timings'][phase+'_ns']/1e6 for x in samples];row[phase+'_p50_ms']=statistics.median(values)
   if phase=='api_sum':row.update(api_p95_ms=quantile(values,.95),api_p99_ms=quantile(values,.99),api_p50_ci95_ms=ci(values,453+i))
  row['process_peak_rss_kib']=int(re.search(r'Maximum resident set size \(kbytes\): (\d+)',raw(f'runs/{i}/resource.log').decode())[1])
  phases={p['label']:p for p in samples[0]['phases']};row['io']={}
  for label in ['opened','planned','published']:
   for owner in ['source','destination']:
    point=phases[label][owner+'_reads']['delta'];value=None if point is None else {k:point[k] for k in ['logical_calls','returned_bytes']};row['io'][label+'_'+owner]=value
    for sample in samples:
     point=next(p for p in sample['phases'] if p['label']==label)[owner+'_reads']['delta'];require((None if point is None else {k:point[k] for k in ['logical_calls','returned_bytes']})==value,'deterministic I/O')
  row['source_work_after_open']=phases['published']['source_cache']['budget_work_used']-phases['opened']['source_cache']['budget_work_used']
  for owner in ['source','destination']:
   for label in ['planned','published','drop_plan']:
    row[owner+'_'+label+'_reserved_bytes']=phases[label][owner+'_budget']['memory_used']
  if lane['instrumentation']=='alloc':
   for phase in ['open_source','open_destination','plan','publication']:
    values=[x['timings'][phase+'_allocation_metrics'] for x in samples]
    for key in ['allocation_calls','allocated_bytes','region_peak_live_bytes','live_bytes_before','live_bytes_after']:
     row[phase+'_'+key]=statistics.median(v[key] for v in values)
    row[phase+'_live_growth']=statistics.median(v['live_bytes_after']-v['live_bytes_before'] for v in values)
    row[phase+'_peak_growth']=statistics.median(v['region_peak_live_bytes']-v['live_bytes_before'] for v in values)
  rows.append(row)
 timing=['open_p50_ms','plan_p50_ms','publication_p50_ms','api_sum_p50_ms','api_p95_ms','api_p99_ms','process_peak_rss_kib']
 allocation=[p+'_'+k for p in ['open_source','open_destination','plan','publication'] for k in ['allocation_calls','allocated_bytes','region_peak_live_bytes','live_growth','peak_growth']]
 pairs=[];flags=[];repeats=[]
 def comparison(a,b,kind):
  metrics=timing if a['instrumentation']=='normal' else allocation
  changes={k:delta(a[k],b[k]) for k in metrics};v={'kind':kind,'from_lane':a['lane'],'to_lane':b['lane'],'instrumentation':a['instrumentation'],'provider':a['provider'],'corpus':a['corpus'],'percent_change':changes}
  for k,change in changes.items():
   if change is None or abs(change)>5:flags.append({**{key:v[key] for key in ['kind','from_lane','to_lane','instrumentation','provider','corpus']},'metric':k,'percent_change':change})
  return v
 for a in rows:
  if a['build']=='baseline':
   b=next(x for x in rows if all(x[k]==a[k] for k in ['provider','corpus','repeat','instrumentation']) and x['build']=='candidate')
   pair=comparison(a,b,'paired');pairs.append(pair)
   require(a['io']==b['io'],'all owner read calls/bytes unchanged');require(a['source_work_after_open']==b['source_work_after_open'],'source work unchanged')
   if a['instrumentation']=='alloc' and a['corpus']=='media-rich':
    require(a['plan_allocated_bytes']-b['plan_allocated_bytes']>=protocol['acceptance']['media_plan_allocated_bytes_reduction_minimum'],'planning allocation gate')
    require(a['plan_live_growth']-b['plan_live_growth']>=protocol['acceptance']['media_plan_live_growth_reduction_minimum'],'planning retained allocation gate')
  if a['repeat']=='R1':
   b=next(x for x in rows if all(x[k]==a[k] for k in ['provider','corpus','build','instrumentation']) and x['repeat']=='R2');repeats.append(comparison(a,b,'repeat'))
 return {'change':453,'normal_samples':480,'allocator_samples':240,'confidence_interval':'2000 bootstrap median resamples seed 453+lane, conditional within each fresh process; linear quantiles; no pooling instrumentation','rows':rows,'pairs':pairs,'repeats':repeats,'review_flags':flags}
def render(v):
 lines=['# PPTX shared decoded payload measurements','','Normal latency and allocator diagnostics use separate binaries and processes.','Each row has 3 warmups / 30 samples. API includes open, plan and publication.','','| Provider | Corpus | Repeat | Baseline API p50 ms | Candidate API p50 ms | Change |','|---|---|---|---:|---:|---:|']
 for pair in v['pairs']:
  a=v['rows'][pair['from_lane']];b=v['rows'][pair['to_lane']]
  if a['instrumentation']=='normal':lines.append(f"| {a['provider']} | {a['corpus']} | {a['repeat']} | {a['api_sum_p50_ms']:.3f} | {b['api_sum_p50_ms']:.3f} | {pair['percent_change']['api_sum_p50_ms']:+.3f}% |")
 lines+=['','| Allocator corpus | Repeat | Plan allocated bytes before / after | Plan live growth before / after | Publication region peak before / after |','|---|---|---:|---:|---:|']
 for pair in v['pairs']:
  a=v['rows'][pair['from_lane']];b=v['rows'][pair['to_lane']]
  if a['instrumentation']=='alloc':lines.append(f"| {a['corpus']} | {a['repeat']} | {a['plan_allocated_bytes']} / {b['plan_allocated_bytes']} | {a['plan_live_growth']} / {b['plan_live_growth']} | {a['publication_region_peak_live_bytes']} / {b['publication_region_peak_live_bytes']} |")
 lines+=['','Raw samples and measurements.json retain phase medians, p95/p99, median intervals,','allocator counters/peaks, process RSS, source work and owner read/budget gauges.','Process RSS includes untimed fixture construction. Allocator region peaks are','absolute live bytes including region entry, not physical RSS. Conservative staging','admission remains. Range is 64 KiB/200 us/25 MiB/s separate-sleep simulation.','All >5% paired and repeat flags are reviewed individually; no native/cold/scaling','coverage is promoted.']
 return '\n'.join(lines)+'\n'
if __name__=='__main__':
 v=derive();(ROOT/'measurements.json').write_text(json.dumps(v,indent=2)+'\n');(ROOT/'measurements.md').write_text(render(v));print(json.dumps({'status':'pass','review_flags':len(v['review_flags'])}))
