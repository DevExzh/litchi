#!/usr/bin/env python3
"""Reduce retained 0940 evidence; run after the independent audit and analysis."""
import hashlib,json,shutil
from pathlib import Path
import sys
r=Path(sys.argv[1]);out=Path(sys.argv[2]);out.mkdir(parents=True,exist_ok=True)
def desc(p):
 b=p.read_bytes();return dict(path=str(p),bytes=len(b),sha256=hashlib.sha256(b).hexdigest())
a=json.loads((r/'descriptive-analysis-v2.json').read_text());old=json.loads((r/'descriptive-analysis-v1.json').read_text())
assert [k for k in old if old[k]!=a[k]]==['independent_review']
allocation=[]
for route in ['owned_vec','owned_reader']:
 rows=[c for c in a['children'] if c['arm']=='allocator_on' and c['case']=='workload' and c['route']==route];assert len(rows)==12
 for phase in ['prep','open']:
  metrics={}
  for key in ['allocated_bytes','allocation_calls','reallocation_calls']:
   ds=[c['allocation'][phase]['metrics'][key] for c in rows];lo=min(d['min'] for d in ds);hi=max(d['max'] for d in ds);assert lo==hi
   metrics[key]=lo
  allocation.append(dict(route=route,phase=phase,samples=600,constant_metrics=metrics))
paths=['descriptive-analysis-v1.json','descriptive-analysis-v2.json','preparation-fault-association-v1.json','source-candidate-manifest-v1.jsonl','manifests/formal-manifest-authorized-v2.json','receipts/independent-formal-audit-v3.json','receipts/independent-formal-audit-terminal-v2.json','receipts/formal-audit-controls-v3.json','receipts/analysis-v2-terminal.json','receipts/fault-association-v1-terminal.json','receipts/source-candidate-v1.json','replay-equivalence-v1.json','target-normal-v1/release/litchi-perf-baseline','target-allocator-v1/release/litchi-perf-baseline-alloc']
res={k:a[k] for k in ['schema','qualification','limits','groups','normal_probe_contrasts','route_contrasts','manifest','independent_review','workload_children','calibration_children','calibration_by_arm_and_boundary']}
res.update(schema='litchi.performance.0940.reduced-evidence.v1',allocation_constants=allocation,analysis_v1_v2_numeric_equality=True,artifacts=[desc(r/p) for p in paths],exploratory_fault_association=json.loads((r/'preparation-fault-association-v1.json').read_text()),preparation_threshold_counts=[{k:c[k] for k in ['id','arm','route','exploratory_preparation_ge_100000ns_count']} for c in a['children'] if c['case']=='workload'])
b=(json.dumps(res,indent=2,sort_keys=True)+'\n').encode();assert len(b)<1048576;(out/'reduced-evidence-v1.json').write_bytes(b)
for p in ['analyze-0940-v1.py','preparation-fault-association-v1.py','docx_owned_attribution_validator_v3.py','independent-formal-packet-checker-v2.py','reduce-evidence-v1.py']:
 shutil.copyfile(r/p,out/p)
print(desc(out/'reduced-evidence-v1.json'))
