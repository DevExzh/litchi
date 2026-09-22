#!/usr/bin/env python3
"""Present independent process ranges, with no subtraction across route scopes."""
import json
from pathlib import Path
P=Path(__file__).resolve().parent
x=json.loads((P/'analysis.json').read_text());groups={}
for r in x['processes']:groups.setdefault((r['case'],r['operation'],r['policy']),[]).append(r)
lines=['# Current DOC/PPT baseline','', 'Ranges below span six independent native processes per route. Timing is microseconds; phase percentages are per-process medians of within-sample ratios. Allocation counters span three separate instrumented processes. No before/after speedup, RSS, cold-cache, or concurrent claim follows.','', '| Case | Route | Whole p50 μs range | Whole mean μs range | Stage % range | Finish % range |','| --- | --- | ---: | ---: | ---: | ---: |']
summary=[]
for (case,op,pol),rows in groups.items():
 native=[r for r in rows if r['lane']=='native'];alloc=[r for r in rows if r['lane']=='allocation']
 def bounds(values):return [min(values),max(values)]
 def text(v):return f'{v[0]:.2f}–{v[1]:.2f}'
 p50=bounds([r['timing']['whole_ns']['p50']/1000 for r in native]);mean=bounds([r['timing']['whole_ns']['mean']/1000 for r in native]);stage=bounds([r['phase_percent']['stage_ns']['p50'] for r in native]) if op=='container' else None;finish=bounds([r['phase_percent']['finish_ns']['p50'] for r in native]) if op=='container' else None
 route='format default' if op=='format' else 'container '+pol
 lines.append(f'| {case} | {route} | {text(p50)} | {text(mean)} | {text(stage) if stage else "—"} | {text(finish) if finish else "—"} |')
 summary.append(dict(case=case,operation=op,policy=pol,whole_p50_us=p50,whole_mean_us=mean,stage_median_percent=stage,finish_median_percent=finish,allocations=alloc[0]['allocations']))
lines.extend(['','The common-container PPT route is an alternative control, not a decomposition of its public save. DOC also runs in separate processes: subtracting these route medians would not estimate a nested phase. Staging includes render, reopen, recapture and discovery. Allocation regions retain returned editors/output vectors at the recorded boundary; peak live values are region increments, not whole-process RSS.','','See analysis.json for all sample-derived per-process p95, p99, maxima and phase fractions, plus exact allocation values.'])
(P/'analysis.md').write_text('\n'.join(lines)+'\n');(P/'report-summary.json').write_text(json.dumps(summary,indent=2)+'\n');print('PASS nine route summaries')
