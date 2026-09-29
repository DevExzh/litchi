"""Descriptive route/cache statistics; never a before/after optimization claim."""
from collections import defaultdict
from pathlib import Path
import argparse,hashlib,json,math,random,statistics

P=Path(__file__).resolve().parent
def read(path):return json.loads(Path(path).read_text())
def digest(path):return hashlib.sha256(Path(path).read_bytes()).hexdigest()
def percentile(values,q):return sorted(values)[max(0,math.ceil(len(values)*q)-1)]
def summary(values):
 assert values
 return dict(min=min(values),p50=percentile(values,.5),p95=percentile(values,.95),p99=percentile(values,.99),max=max(values),mean=statistics.mean(values))
def spread(values):
 return dict(min=min(values),median=statistics.median(values),max=max(values),max_over_min=max(values)/min(values) if min(values)>0 else None)
def compute():
 validation=read(P/'capture-admission.json');assert validation['status']=='pass'
 for name,descriptor in validation['inputs'].items():
  assert digest(P/name)==descriptor['sha256']
 capture=read(P/'capture.json');assert capture['status']=='commands_pass'
 groups=defaultdict(list)
 for item in capture['rows']:
  plan=item['plan'];d=item['result']['report'];path=P/Path(d['path']).name
  assert digest(path)==d['sha256'] and path.stat().st_size==d['bytes']
  report=read(path);assert len(report['results'])==len(report['filesystem_evidence'])==1
  result=report['results'][0];evidence=report['filesystem_evidence'][0]
  samples=evidence['samples'];times=[s['elapsed_ns'] for s in samples]
  assert times==result['elapsed_ns']['samples'] and len(times)==30
  assert result['case']==plan['case'] and result['cache_state']==plan['cache_state']
  process={k:summary([s['process_metrics'][k] for s in samples]) for k in samples[0]['process_metrics'] if k!='clock_ticks_per_second'}
  groups[(plan['case'],plan['cache_state'])].append(dict(block=plan['block'],latency_ns=summary(times),process_metrics=process,report_sha256=d['sha256']))
 assert len(groups)==12 and all(len(rows)==6 for rows in groups.values())
 distributions=[];lookup={};flags=[]
 for (case,state),rows in sorted(groups.items()):
  rows.sort(key=lambda row:row['block']);assert [r['block'] for r in rows]==list(range(6))
  latency={k:spread([r['latency_ns'][k] for r in rows]) for k in rows[0]['latency_ns']}
  process={k:{q:spread([r['process_metrics'][k][q] for r in rows]) for q in rows[0]['process_metrics'][k]} for k in rows[0]['process_metrics']}
  for metric in ['p50','p95','p99']:
   if latency[metric]['max_over_min']>1.2:flags.append(dict(case=case,state=state,metric=metric,reason='block spread exceeds 20%',spread=latency[metric]['max_over_min']))
  row=dict(case=case,cache_state=state,blocks=rows,latency_ns=latency,process_metrics=process)
  distributions.append(row);lookup[(case,state)]=row
 pairs=[]
 for eager,source in [('opc_file_eager_open','opc_file_source_open'),('opc_file_eager_one_part_atomic_save','opc_file_source_one_part_atomic_save'),('pptx_file_eager_open_selected_slide_lifecycle','pptx_file_source_open_selected_slide_lifecycle')]:
  for state in ['warm','cold-verified']:
   ratios=[e['latency_ns']['p50']/s['latency_ns']['p50'] for e,s in zip(lookup[(eager,state)]['blocks'],lookup[(source,state)]['blocks'])]
   rng=random.Random(834083);boot=sorted(statistics.median(rng.choices(ratios,k=6)) for _ in range(10000))
   pairs.append(dict(eager=eager,source=source,cache_state=state,per_block_eager_over_source_p50=ratios,median=statistics.median(ratios),bootstrap_95=[boot[249],boot[9749]],scope='descriptive route ratio; not an optimization effect'))
 return dict(schema='litchi.0834.descriptive-analysis.v1',reports=72,samples=2160,bootstrap_seed=834083,bootstrap_resamples=10000,distributions=distributions,route_ratios=pairs,spread_flags=flags,limitations=['Synthetic fixed corpora and one host only.','OPC eager open drops its package inside the timer; source open retains it through post-timer diagnostics.','PPTX logical source counters come from an untimed replay.','Verified cold means observed page-cache residency and process read_bytes, not physical-device I/O.','Route/cache baseline only; no production optimization or before/after claim.'])
def main():
 parser=argparse.ArgumentParser();parser.add_argument('--write',action='store_true');parser.add_argument('--check',action='store_true');args=parser.parse_args()
 result=compute();data=json.dumps(result,indent=2,sort_keys=True)+'\n';path=P/'analysis.json'
 if args.write:
  with path.open('x') as f:f.write(data)
 if args.check:assert path.read_text()==data
 print(json.dumps(dict(reports=result['reports'],samples=result['samples'],route_ratios=result['route_ratios'],spread_flags=len(result['spread_flags'])),indent=2))
if __name__=='__main__':main()
