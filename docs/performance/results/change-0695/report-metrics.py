#!/usr/bin/env python3
"""Derive scoped attribution ratios and all native table rows."""
from pathlib import Path
import json
P=Path(__file__).resolve().parent
s={r['case']:r for r in json.loads((P/'summary.json').read_text())}
rows=['| Case | Calls per sequence | Input bytes | Median range (µs) |','| --- | ---: | ---: | ---: |']
for name in ['all','presentation','slides']+['slide'+str(i) for i in range(1,14)]:
 seq=P/'sequences'/f'{name}.txt';sources=[(seq.parent/line).resolve() for line in seq.read_text().splitlines()]
 r=s[name]
 rows.append(f"| {name} | {len(sources)} | {sum(f.stat().st_size for f in sources):,} | {r['leg_median_min_ns']/1000:.3f}–{r['leg_median_max_ns']/1000:.3f} |")
(P/'tables.md').write_text('\n'.join(rows)+'\n')
total=s['all']['median_of_leg_medians_ns']
ratios={name:s[name]['median_of_leg_medians_ns']/total for name in ['presentation','slides','slide11']}
ratios['sum_individual_over_full']=(s['presentation']['median_of_leg_medians_ns']+sum(s['slide'+str(i)]['median_of_leg_medians_ns'] for i in range(1,14)))/total
ratios['sum_groups_over_full']=(s['presentation']['median_of_leg_medians_ns']+s['slides']['median_of_leg_medians_ns'])/total
(P/'attribution.json').write_text(json.dumps(dict(denominator='median of four leg medians of isolated all-sequence time; not capture latency',ratios=ratios,max_leg_median_spread_pct=max(r['leg_median_spread_pct'] for r in s.values())),indent=2)+'\n')
