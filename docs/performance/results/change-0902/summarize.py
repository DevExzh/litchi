"""Descriptive tables from the admitted raw population, separate from gate reduction."""
import json,pathlib,statistics
p=pathlib.Path(__file__).resolve().parent;a=json.loads((p/'analysis.json').read_text())
med=statistics.median
cases={}
for case in a['reduction']['cases']:
 out={};cases[case]=out
 for role in ['baseline','candidate']:
  rows=[r for r in a['native'] if r['case_id']==case and r['role']==role]
  values={metric:[r[metric] for r in rows] for metric in ['p50','p95','p99']}
  desc={'native':{metric:{'median':med(v),'min':min(v),'max':max(v)} for metric,v in values.items()},'p50_spread':max(values['p50'])/min(values['p50']),'rss_kib':{'median':med(r['resource']['max_rss_kib'] for r in rows),'min':min(r['resource']['max_rss_kib'] for r in rows),'max':max(r['resource']['max_rss_kib'] for r in rows)}}
  phase_medians={}
  for row in rows:
   raw=json.loads(pathlib.Path(row['report']['path']).read_text())['results'][0]
   source=(raw.get('source') or {}).get('xlsx_cell_values',{})
   if source:
    total=[sum(source[k][i] for k in ['open_ns','plan_ns','commit_ns','publication_ns']) for i in range(row['samples'])]
    assert total==raw['elapsed_ns']['samples']
    for phase in ['open_ns','plan_ns','commit_ns','publication_ns','reopen_ns']:
     phase_medians.setdefault(phase,[]).append(med(source[phase]))
  desc['phase_p50_ns']={k:{'median':med(v),'min':min(v),'max':max(v)} for k,v in phase_medians.items()}
  allocrows=[r for r in a['allocator'] if r['case_id']==case and r['role']==role];metrics={}
  for row in allocrows:
   for phase,vector in row['allocator_metrics'].items():
    for sample in vector:
     for key,value in sample.items():
      if isinstance(value,int):metrics.setdefault(phase,{}).setdefault(key,[]).append(value)
     if sample.get('status')=='measured':metrics.setdefault(phase,{}).setdefault('region_peak_above_entry_live_bytes',[]).append(sample['region_peak_live_bytes']-sample['live_bytes_before'])
  desc['phase_allocations']={phase:{key:{'median':med(v),'min':min(v),'max':max(v)} for key,v in fields.items()} for phase,fields in metrics.items()}
  desc['allocator_metrics_status']=sorted({r['allocator_metrics_status'] for r in allocrows})
  out[role]=desc
 out['ratios_of_native_medians']={metric:out['candidate']['native'][metric]['median']/out['baseline']['native'][metric]['median'] for metric in ['p50','p95','p99']}
result={'status':'pass','scope':'descriptive medians/ranges; frozen acceptance remains analysis.reduction paired ratios','cases':cases}
with (p/'summary.json').open('x') as f:json.dump(result,f,indent=2,sort_keys=True);f.write('\n')
for case,out in cases.items():
 print(case,'tails',out['ratios_of_native_medians'],'rss',out['baseline']['rss_kib'],out['candidate']['rss_kib'])
 for role in ['baseline','candidate']:
  q=out[role]['phase_allocations'].get('publication_allocation_metrics',{});print(role,'publication', {k:v for k,v in q.items() if k in ['allocation_calls','allocated_bytes','reallocation_calls','region_peak_above_entry_live_bytes','region_peak_live_bytes']})
