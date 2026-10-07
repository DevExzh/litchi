#!/usr/bin/env python3
"""Summarize already sealed 0961 role-A native timings; launch no workload."""
import argparse, hashlib, json, statistics
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--packet',type=Path,required=True);p.add_argument('--plan',type=Path,required=True);p.add_argument('--output',type=Path,required=True);a=p.parse_args()
plan=json.loads(a.plan.read_text());groups={};bindings=[]
for child in plan['lanes']['native']['children']:
 if child['binary_role']!='baseline':continue
 d=a.packet/'children'/child['child_id'];path=d/'output/report.json';data=path.read_bytes();report=json.loads(data)
 assert report['status']=='completed' and report['workers']==1 and report['measured_samples']==60
 assert (d/'validation.json').is_file() and (d/'terminal.json').is_file()
 samples=report['case']['samples'];assert len(samples)==60
 assert all(s['total_ns']>=sum(s[k] for k in ['ingress_ns','edit_commit_ns','publication_ns','commit_drop_ns']) for s in samples)
 bindings.append({'child_id':child['child_id'],'report':str(path),'bytes':len(data),'sha256':hashlib.sha256(data).hexdigest()})
 groups.setdefault(child['case_id'],[]).append(samples)
result={'schema':'litchi.performance.0962.retained-native-phase-summary.v1','source_plan':str(a.plan),'source_plan_sha256':hashlib.sha256(a.plan.read_bytes()).hexdigest(),'scope':'0961 baseline-only unprofiled native samples; ratios of phase sums to total sums per process. Not a new benchmark or causal speedup prediction.','bindings':bindings,'cases':{}}
for case,children in sorted(groups.items()):
 assert len(children)==6
 rows=[r for samples in children for r in samples]
 phases={}
 for key in ['ingress_ns','edit_commit_ns','publication_ns','commit_drop_ns']:
  fractions=[sum(r[key] for r in samples)/sum(r['total_ns'] for r in samples) for samples in children]
  phases[key]={'median_process_fraction':statistics.median(fractions),'min_process_fraction':min(fractions),'max_process_fraction':max(fractions),'median_ns_across_360_samples':statistics.median(r[key] for r in rows)}
 f=phases['edit_commit_ns']['median_process_fraction']
 result['cases'][case]={'processes':6,'samples':len(rows),'phases':phases,'total_median_ns_across_360_samples':statistics.median(r['total_ns'] for r in rows),'ideal_edit_phase_removal_ceiling':1/(1-f),'ideal_edit_phase_10x_total_speedup':1/(1-f+f/10),'model_caveat':'Amdahl model assumes other phases unchanged and no overlap; fraction is workload-specific, not measured acceleration.'}
with a.output.open('x') as f:json.dump(result,f,indent=2);f.write('\n')
for case,v in result['cases'].items():print(case,round(v['phases']['edit_commit_ns']['median_process_fraction'],4),round(v['ideal_edit_phase_removal_ceiling'],3))
