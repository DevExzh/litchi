#!/usr/bin/env python3
"""Expose every native/repeat central, tail and drift comparison without filtering samples."""
import json
from pathlib import Path
P=Path(__file__).resolve().parent
x=json.loads((P/'analysis.json').read_text())
rows=[]; failures=[]
for kind in ('native','repeat'):
 for r in x[kind]:
  metrics=r['timing'] if kind=='native' else {'loop-mean':r['timing']}
  for metric,legs in metrics.items():
   for pair,a,b in [('aa2/aa1','aa1','aa2'),('b1/a1','a1','b1'),('b2/a2','a2','b2'),('a2/a1','a1','a2'),('b2/b1','b1','b2')]:
    for stat in ('p50','mean','p95','p99','maximum'):
     av,bv=legs[a][stat],legs[b][stat]
     rows.append(dict(kind=kind,case=r['case'],mode=r['mode'],metric=metric,pair=pair,statistic=stat,baseline=av,candidate=bv,percent=(bv/av-1)*100 if av else None))
  pairs=r['paired'] if kind=='native' else {'loop-mean':r['paired']}
  for metric,pairmap in pairs.items():
   for pair,g in pairmap.items():
    for stat in ('p50','mean'):
     if not g[stat]['pass']:failures.append(dict(kind=kind,case=r['case'],mode=r['mode'],metric=metric,pair=pair,statistic=stat,**g[stat]))
summary={
 'failed_native_groups':sum(not r['timing_gate_pass'] for r in x['native']),
 'failed_repeat_groups':sum(not r['timing_gate_pass'] for r in x['repeat']),
 'failed_central_statistics':len(failures),
 'paired_tail_regression_flags':sum(r['pair'] in ('b1/a1','b2/a2') and r['statistic'] in ('p95','p99','maximum') and r['percent']>5 for r in rows),
 'aa_central_drift_flags':sum(r['pair']=='aa2/aa1' and r['statistic'] in ('p50','mean') and abs(r['percent'])>5 for r in rows),
 'repeat_central_drift_flags':sum(r['pair'] in ('a2/a1','b2/b1') and r['statistic'] in ('p50','mean') and abs(r['percent'])>5 for r in rows),
}
(P/'measurement-details.json').write_text(json.dumps(dict(summary=summary,failed_gates=failures,comparisons=rows),indent=2)+'\n')
lines=['# Complete measurement comparisons','', 'Negative latency deltas are faster. Repeat statistics describe nine process loop means, not individual-query tails. All native tails, A/A and within-phase drift are diagnostic; frozen central gates decide rejection.','',json.dumps(summary,sort_keys=True),'','## Failed frozen central checks','','| Lane | Case | Mode | Metric | Pair | Statistic | Delta |','| --- | --- | --- | --- | --- | --- | ---: |']
for r in failures:lines.append(f"| {r['kind']} | {r['case']} | {r['mode']} | {r['metric']} | {r['pair']} | {r['statistic']} | {r['percent']:+.2f}% |")
lines+=['','## All comparisons','','| Lane | Case | Mode | Metric | Pair | Statistic | A ns | B ns | Delta |','| --- | --- | --- | --- | --- | --- | ---: | ---: | ---: |']
for r in rows:lines.append(f"| {r['kind']} | {r['case']} | {r['mode']} | {r['metric']} | {r['pair']} | {r['statistic']} | {r['baseline']:.3f} | {r['candidate']:.3f} | {r['percent']:+.2f}% |")
(P/'measurements.md').write_text('\n'.join(lines)+'\n')
print(json.dumps(summary))
