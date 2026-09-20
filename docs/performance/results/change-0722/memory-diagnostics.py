#!/usr/bin/env python3
"""Recompute aligned allocation-region diagnostics; no cross-phase sums."""
import hashlib,json,sys
from pathlib import Path
P=Path(__file__).resolve().parent
rows=[];inputs={}
for stage in ['baseline-A1','candidate-B1','candidate-B2','baseline-A2']:
 for corpus in ['generated','numbered-list']:
  for phase in ['edit','lifecycle']:
   f=P/f'{stage}-allocator-{corpus}-{phase}.json';inputs[f.name]=hashlib.sha256(f.read_bytes()).hexdigest()
   a=json.loads(f.read_text())['results'][0]['operation_metrics']['allocation'];assert a['status']=='measured'
   values={key:a[key]['values'] for key in ['live_bytes_before','live_bytes_after','region_peak_live_bytes']};assert all(len(v)==3 for v in values.values())
   net=[x-y for x,y in zip(values['live_bytes_after'],values['live_bytes_before'])]
   peak=[x-y for x,y in zip(values['region_peak_live_bytes'],values['live_bytes_before'])]
   rows.append({'stage':stage,'corpus':corpus,'phase':phase,'net_live_bytes':net,'peak_above_start_bytes':peak})
result={'scope':'Aligned per-operation region counters; owner destruction outside region, not RSS or leak evidence; phases not additive','formulas':{'net_live_bytes':'live_bytes_after - live_bytes_before','peak_above_start_bytes':'region_peak_live_bytes - live_bytes_before'},'inputs':inputs,'script_sha256':hashlib.sha256(Path(__file__).read_bytes()).hexdigest(),'rows':rows}
encoded=json.dumps(result,indent=2)+'\n';out=P/'memory-diagnostics.json'
if sys.argv[1:]==['--check']:assert out.read_text()==encoded
else:assert sys.argv[1:]==[];out.write_text(encoded)
print('Verified 16 aligned memory diagnostic reports')
