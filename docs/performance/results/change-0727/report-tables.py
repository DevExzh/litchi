#!/usr/bin/env python3
"""Summarize recurrence without pooling owners or discarding any observation."""
import json
from pathlib import Path
P=Path(__file__).resolve().parent
x=json.loads((P/'analysis.json').read_text());plan=json.loads((P/'plan.json').read_text());rows=[]
for c in plan['cases']:
 for metric in plan['timing_metrics']:
  r=[r for r in x['comparisons'] if r['case']==c['id'] and r['metric']==metric];paired=[v for v in r if v['pair'] in ['b1/a1','b2/a2']];assert len(paired)==18
  row=dict(case=c['id'],metric=metric,role=c['role'],focus=metric==c['focus_metric'],pairs=len(paired),failed_p50=sum(not v['statistics']['p50']['pass_'] for v in paired),failed_mean=sum(not v['statistics']['mean']['pass_'] for v in paired),p50_percent_range=[min(v['statistics']['p50']['percent'] for v in paired),max(v['statistics']['p50']['percent'] for v in paired)],mean_percent_range=[min(v['statistics']['mean']['percent'] for v in paired),max(v['statistics']['mean']['percent'] for v in paired)],aa_mean_flags=sum(abs(v['statistics']['mean']['percent'])>5 for v in r if v['pair']=='aa2/aa1'),within_mean_flags=sum(abs(v['statistics']['mean']['percent'])>5 for v in r if v['pair'] in ['a2/a1','b2/b1']))
  rows.append(row)
(P/'report-summary.json').write_text(json.dumps(dict(summary=x['summary'],rows=rows),indent=2)+'\n')
s=['# 0727 complete diagnostic comparisons','','No retention gate is passed by this diagnostic. 0726 remains rejected. Each comparison pairs independent processes at the same cycle and replicate ordinal; every process retains 100 owners. Ranges are descriptive, not confidence intervals.','','## Every cell and metric','','| Cell | Metric | Pairs | p50 failures | Mean failures | p50 delta range | Mean delta range | A/A mean drift flags | Within-phase mean drift flags |','| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |']
for r in rows:s.append(f"| {r['case']} | {r['metric']} | {r['pairs']} | {r['failed_p50']} | {r['failed_mean']} | {r['p50_percent_range'][0]:+.2f}% to {r['p50_percent_range'][1]:+.2f}% | {r['mean_percent_range'][0]:+.2f}% to {r['mean_percent_range'][1]:+.2f}% | {r['aa_mean_flags']} | {r['within_mean_flags']} |")
s+=['','## Every failed paired central check','','| Cycle | Replicate | Cell | Metric | Pair | Statistic | A ns | B ns | Delta |','| ---: | ---: | --- | --- | --- | --- | ---: | ---: | ---: |']
for r in x['failed_checks']:s.append(f"| {r['cycle']} | {r['replicate']} | {r['case']} | {r['metric']} | {r['pair']} | {r['statistic']} | {r['baseline']:.3f} | {r['candidate']:.3f} | {r['percent']:+.3f}% |")
(P/'analysis.md').write_text('\n'.join(s)+'\n');print(json.dumps(x['summary']))
