#!/usr/bin/env python3
"""Retain phase diagnostics for adverse lifecycle cases; no causal inference."""
import csv,json
from pathlib import Path
HERE=Path(__file__).resolve().parent
result=[]
for stem in ['p512-k8-owned-repeated','p512-k8-file-repeated','p128-k1-owned-repeated','p128-k1-owned-batch']:
    for phase in ['before','after']:
        path=HERE/phase/(stem+'.csv')
        if not path.exists():continue
        rows=[r for r in csv.DictReader(path.open()) if r['warmup']=='false']
        worst=max(rows,key=lambda r:int(r['elapsed_ns']))
        result.append({'phase':phase,'workload':stem,'worst_sample':{key:int(worst[key]) for key in ['repeat','ordinal','elapsed_ns','open_ns','edit_ns','commit_ns','publish_ns','drop_ns']}})
(HERE/'adverse-diagnostics.json').write_text(json.dumps({'scope':'Worst individual samples identify observed phases; phase percentiles need not occur in the same sample. Shared-host order and whole-child RSS prevent isolated causal claims. No samples discarded or formal captures replaced.','cases':result},indent=2)+'\n')
for row in result:print(row)
