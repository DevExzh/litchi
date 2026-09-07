#!/usr/bin/env python3
"""Derive operation growth and regional peak above entry from retained counters."""
import json,statistics
from pathlib import Path
ROOT=Path(__file__).resolve().parent
p=json.loads((ROOT/'protocol.json').read_text());rows=[]
for i,lane in enumerate(p['lanes']):
    if lane['instrumentation']!='allocator':continue
    report=json.loads((ROOT/'runs'/str(i)/'report.json').read_text());phases={}
    for phase in ['open_source','open_destination','plan','publication']:
        values={k:[] for k in ['allocation_calls','allocated_bytes','live_growth_bytes','regional_peak_above_entry_bytes']}
        for sample in report['samples_raw']:
            m=sample['timings'][phase+'_allocation_metrics'];values['allocation_calls'].append(m['allocation_calls']);values['allocated_bytes'].append(m['allocated_bytes']);values['live_growth_bytes'].append(m['live_bytes_after']-m['live_bytes_before']);values['regional_peak_above_entry_bytes'].append(m['region_peak_live_bytes']-m['live_bytes_before'])
        phases[phase]={k:{'min':min(v),'median':statistics.median(v),'max':max(v)} for k,v in values.items()}
    rows.append({'lane':i,**lane,'phases':phases})
result={'scope':'Per-operation heap counters; regional peak above entry subtracts live_bytes_before from region_peak_live_bytes for each sample. The existing 64 KiB stack buffer is unchanged and outside allocator accounting. Absolute process live/peak values can have different pre-operation baselines.','rows':rows}
(ROOT/'allocation-summary.json').write_text(json.dumps(result,indent=2)+'\n');print(json.dumps({'status':'pass','lanes':len(rows)}))
