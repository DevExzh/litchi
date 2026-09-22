#!/usr/bin/env python3
"""Render per-process ranges and observer-control flags without scope subtraction."""
import json
from pathlib import Path
P=Path(__file__).resolve().parent;x=json.loads((P/'analysis.json').read_text())
def bounds(v):return [min(v),max(v)]
def fmt(v):return f'{v[0]:.2f}–{v[1]:.2f}'
lines=['# Public DOC attribution','', 'All ranges span nine independent process statistics per case/route. Whole timing is microseconds. Fractions are computed within each measured owner and then summarized per process. No cross-route median is subtracted to invent a nested phase.','', '| Case | Route | Whole p50 μs | Whole mean μs |','| --- | --- | ---: | ---: |'];summary=[]
for case in ['docfloat','docnohf']:
 for route in ['ordinary-opaque','ordinary-split','profiled-empty','profiled-clock']:
  rows=[r for r in x['processes'] if r['case']==case and r['route']==route]
  entry=dict(case=case,route=route,whole_p50_us=bounds([r['timing']['whole_ns']['p50']/1000 for r in rows]),whole_mean_us=bounds([r['timing']['whole_ns']['mean']/1000 for r in rows]),whole_percent={k:bounds([r['whole_percent'][k]['p50'] for r in rows]) for k in rows[0]['whole_percent']},parent_percent={k:bounds([r['parent_percent'][k]['p50'] for r in rows]) for k in rows[0]['parent_percent']})
  summary.append(entry);lines.append(f"| {case} | {route} | {fmt(entry['whole_p50_us'])} | {fmt(entry['whole_mean_us'])} |")
lines.extend(['','## Outer phases, ordinary split','','| Case | Phase | Median % of same owner whole |','| --- | --- | ---: |'])
for e in summary:
 if e['route']=='ordinary-split':
  for k,v in e['whole_percent'].items():lines.append(f"| {e['case']} | {k} | {fmt(v)} |")
lines.extend(['','## Semantic phases, profiled clock','','| Case | Phase | Median % of same owner whole | Median % of parent phase |','| --- | --- | ---: | ---: |'])
for e in summary:
 if e['route']=='profiled-clock':
  for k,v in e['parent_percent'].items():lines.append(f"| {e['case']} | {k} | {fmt(e['whole_percent'][k])} | {fmt(v)} |")
lines.extend(['','## Matched observer controls','','Positive deltas are slower. Flags count absolute changes above 5% in either direction; they are interpretation flags, not optimization admission gates. Each pair matches case, cycle and round.','','| Case | Before → after | p50 delta % | p50 flags / 9 | Mean delta % | Mean flags / 9 |','| --- | --- | ---: | ---: | ---: | ---: |'])
controls=[]
for case in ['docfloat','docnohf']:
 for before,after in [('ordinary-opaque','ordinary-split'),('ordinary-split','profiled-empty'),('profiled-empty','profiled-clock')]:
  rows=[r for r in x['comparisons'] if r['case']==case and r['before']==before and r['after']==after];entry=dict(case=case,before=before,after=after)
  for metric in ['p50','mean']:entry[metric]=dict(delta_range=bounds([r['metrics'][metric]['delta_percent'] for r in rows]),flags=sum(r['metrics'][metric]['flag'] for r in rows))
  controls.append(entry);lines.append(f"| {case} | {before} → {after} | {fmt(entry['p50']['delta_range'])} | {entry['p50']['flags']} | {fmt(entry['mean']['delta_range'])} | {entry['mean']['flags']} |")
lines.extend(['','Profiled open releases its strict editor before subsequent validation, unlike ordinary open. The profiled implementation comparison therefore includes lifetime/code-path changes. Timestamp recorder calibration is measured outside the workflow and is never subtracted from workflow timing. This build enables performance-diagnostics for every route; the experiment does not measure feature-enabled versus default-build code generation.'])
(P/'analysis.md').write_text('\n'.join(lines)+'\n');(P/'report-summary.json').write_text(json.dumps(dict(routes=summary,controls=controls),indent=2)+'\n');print('PASS eight route summaries and six observer comparisons')
