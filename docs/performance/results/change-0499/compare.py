#!/usr/bin/env python3
"""Matched before/after comparisons with raw-oracle and threshold verification."""
import csv,hashlib,json,math,re
from collections import Counter
from pathlib import Path
HERE=Path(__file__).resolve().parent
RSS=re.compile(r'Maximum resident set size \(kbytes\):\s*(\d+)')

def load(phase,name):
 p=HERE/phase/name
 receipt=json.loads(p.with_suffix('.json').read_text())
 assert receipt['exit_code']==0 and receipt['cleanup_verified'],str(p)
 assert hashlib.sha256(p.with_suffix('.csv').read_bytes()).hexdigest()==receipt['csv_sha256']
 assert hashlib.sha256(p.with_suffix('.time.stderr').read_bytes()).hexdigest()==receipt['stderr_sha256']
 rows=[r for r in csv.DictReader(p.with_suffix('.csv').open()) if r['record']=='sample']
 assert {(int(r['repeat']),int(r['index'])) for r in rows}=={(a,b) for a in [0,1] for b in range(33)}
 assert len(rows)==66
 assert all(int(r['budget_memory_after_release'])==0 and int(r['budget_objects_after_release'])==0 for r in rows)
 oraclekeys=['corpus_sha256','logical_bytes','digest','source_calls','source_requested_bytes','source_returned_bytes','cache_cold_loads','cache_hits','budget_work','budget_input_bytes']
 signatures={tuple(r[k] for k in oraclekeys) for r in rows};assert len(signatures)==1
 rss=int(RSS.search(p.with_suffix('.time.stderr').read_text()).group(1))
 return [r for r in rows if int(r['index'])>=3],rss,signatures

def stats(rows):
 ns=sorted(int(r['elapsed_ns']) for r in rows)
 values={f'p{p}_us':ns[math.ceil(len(ns)*p/100)-1]/1000 for p in [50,95,99]}
 values['mean_us']=sum(ns)/len(ns)/1000
 values['throughput_bytes_s']=sum(int(r['logical_bytes']) for r in rows)*1e9/sum(ns)
 return values

def compare(before,after):
 result={};flags=[]
 for metric,value in before.items():
  delta=(after[metric]/value-1)*100
  result[metric]={'before':value,'after':after[metric],'delta_pct':delta}
  if delta>5 and metric!='throughput_bytes_s' or delta < -5 and metric=='throughput_bytes_s': flags.append(metric)
 return {'metrics':result,'adverse_flags':flags}

comparisons=[]
for corpus in ['few-large','many-small']:
 for source in ['owned','file','instrumented']:
  for api,workers in [('serial',1),('batch',1),('batch',2),('batch',4),('batch',8)]:
   name=f'{corpus}-{source}-{api}-w{workers}'
   before,brss,boracle=load('before',name);after,arss,aoracle=load('after',name)
   assert boracle==aoracle,name
   b=stats(before);a=stats(after);b['rss_kib']=brss;a['rss_kib']=arss
   aggregate=compare(b,a)
   repeats={str(i):compare(stats([r for r in before if int(r['repeat'])==i]),stats([r for r in after if int(r['repeat'])==i])) for i in [0,1]}
   comparisons.append({'name':name,'corpus':corpus,'source':source,'api':api,'workers':workers,'aggregate':aggregate,'repeats':repeats})
counts=Counter(flag for c in comparisons for flag in c['aggregate']['adverse_flags'])
repeatcounts=Counter(flag for c in comparisons for r in c['repeats'].values() for flag in r['adverse_flags'])
result={'scope':'matched unchanged-harness before/after; RSS once per whole child, no repeated RSS observations',
        'children':60,'measured_samples':3600,'warmup_samples':360,'comparisons':comparisons,
        'aggregate_flag_counts':dict(counts),'repeat_flag_counts':dict(repeatcounts),
        'byte_counter_budget_oracles_equal':True}
(HERE/'comparison.json').write_text(json.dumps(result,indent=2)+'\n')
lines=['# 0499 matched before/after measurements','',
 'Sixty children contain 3,600 measured samples and 360 warmups. All byte, source-counter, work-counter, and released-budget oracles match. Percentiles use nearest rank. Whole-child RSS includes setup and verification and is counted once per capture. Positive latency/RSS deltas and negative throughput deltas greater than five percent are retained as adverse flags.','',
 '| Workload | p50 before µs | p50 after µs | p50 change | p99 change | RSS change | Adverse flags |',
 '| --- | ---: | ---: | ---: | ---: | ---: | --- |']
for c in comparisons:
 m=c['aggregate']['metrics'];flags=', '.join(c['aggregate']['adverse_flags']) or '—'
 lines.append(f"| {c['name']} | {m['p50_us']['before']:.2f} | {m['p50_us']['after']:.2f} | {m['p50_us']['delta_pct']:+.2f}% | {m['p99_us']['delta_pct']:+.2f}% | {m['rss_kib']['delta_pct']:+.2f}% | {flags} |")
lines+=['','Aggregate flag counts: '+json.dumps(dict(counts),sort_keys=True)+'.',
        'Per-repeat latency/throughput flag counts: '+json.dumps(dict(repeatcounts),sort_keys=True)+'.',
        '', 'Per-repeat metrics and throughput deltas are retained in comparison.json. Shared-host variation and sparse tail samples limit causal interpretation; no broad end-to-end or isolated-host result is implied.']
(HERE/'comparison.md').write_text('\n'.join(lines)+'\n')
print(json.dumps({'comparisons':len(comparisons),'aggregate_flags':sum(counts.values()),'repeat_flags':sum(repeatcounts.values())}))
