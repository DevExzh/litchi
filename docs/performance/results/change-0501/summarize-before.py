#!/usr/bin/env python3
"""Summarize fresh control API intervals from retained samples."""
from pathlib import Path
import json, statistics
HERE=Path(__file__).resolve().parent
rows=[]
for path in sorted((HERE/'before').glob('*.report.json')):
    report=json.loads(path.read_text()); samples=report['samples_raw']
    rows.append({'name':path.name,'samples':len(samples),**{key+'_p50_ms':statistics.median(s['timings'][field] for s in samples)/1e6 for key,field in [('api','api_sum_ns'),('plan','plan_ns'),('publication','publication_ns')]}})
(HERE/'before-summary.json').write_text(json.dumps({'scope':'Fresh before measurements; medians over recorded30samples; no targeted CPU attribution','rows':rows},indent=2)+'\n')
print(json.dumps({'reports':len(rows),'samples':sum(row['samples'] for row in rows)}))
