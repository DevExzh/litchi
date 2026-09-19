#!/usr/bin/env python3
"""Produce review tables without averaging away individual regressions."""
import json
from pathlib import Path
P=Path(__file__).resolve().parent
native=json.loads((P/'native-comparison.json').read_text())
lines=['# Native query and workflow medians','','Units are ns; paired changes compare B1/A1 and B2/A2. All means, tails,','bootstrap intervals and control drift remain in `native-comparison.json`.','','| Case | Source | Phase | A1 | B1 | B2 | A2 | Paired change % | A/A % |','|---|---|---|---:|---:|---:|---:|---:|---:|']
reg=['# Native paired regression review triggers','','Every row below exceeds 5% p50 in both candidate/baseline pairs.','Single-leg changes and all tail metrics remain in the full comparison.','','| Case | Source | Phase | A1 | B1 | B2 | A2 | Paired change % |','|---|---|---|---:|---:|---:|---:|---:|']
for r in native:
 t=r['timing'];v=r['percent_p50'];row='| '+r['case']+' | '+r['mode']+' | '+r['metric']+' | '+' | '.join(f"{t[k]['p50']:g}" for k in ['a1','b1','b2','a2'])+f" | {v['b1_a1']:+.2f} / {v['b2_a2']:+.2f} |"
 if r['metric'] in ['q8','open-plus-eight']:lines.append(row+f" {v['aa']:+.2f} |")
 if r['review_regression']:reg.append(row)
(P/'native-summary.md').write_text('\n'.join(lines)+'\n');(P/'regressions.md').write_text('\n'.join(reg)+'\n')
lines=['# Allocation and retention changes','','Separate instrumented probes, three identical repeats per group. Bytes are','allocator requests/live gauges, not process RSS or logical-budget guarantees.','All phases, including unchanged rows, remain in `allocation-comparison.json`.','','| Case | Source | Query | Calls before → after | Requested bytes before → after | Peak live before → after | Retained delta before → after |','|---|---|---|---:|---:|---:|---:|']
for r in json.loads((P/'allocation-comparison.json').read_text()):
 b=r['baseline'];a=r['candidate'];keys=['allocation_calls','allocated_bytes','peak_live_delta','retained_live_delta']
 if any(b[k]!=a[k] for k in keys):lines.append('| '+r['case']+' | '+r['mode']+' | '+r['operation']+' | '+' | '.join(f'{b[k]:,} → {a[k]:,}' for k in keys)+' |')
if sum(line.startswith('| ') for line in lines)==1:lines.insert(5,'All measured allocation and retention fields are unchanged in all 96 groups.')
(P/'allocation-summary.md').write_text('\n'.join(lines)+'\n')
lines=['# Additional mean and tail review triggers','','Both paired changes exceed 5%. These are descriptive distributions from 100','samples per leg; individual tail samples and control drift can be noisy.','They are retained even when the median improves. Full distributions and A/A','controls remain in `native-comparison.json`.','','| Case | Source | Phase | Statistic | B1/A1 % | B2/A2 % | A/A % |','|---|---|---|---|---:|---:|---:|']
for r in native:
 for stat in ['mean','p95','p99']:
  t=r['timing'];a=(t['b1'][stat]/t['a1'][stat]-1)*100;b=(t['b2'][stat]/t['a2'][stat]-1)*100;aa=(t['aa2'][stat]/t['aa1'][stat]-1)*100
  if a>5 and b>5:lines.append(f"| {r['case']} | {r['mode']} | {r['metric']} | {stat} | {a:+.2f} | {b:+.2f} | {aa:+.2f} |")
(P/'tail-regressions.md').write_text('\n'.join(lines)+'\n')
