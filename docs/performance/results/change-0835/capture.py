"""Serial formal capture, gated by the retained independent admission."""
import json
import driver as d

def main():
 p=d.P
 admission=d.read(p/'admission.json')
 assert admission['status']=='pass'
 for name,descriptor in admission['inputs'].items():
  assert d.desc(p/name)==descriptor,name
 assert d.read(p/'quality-reuse.json')['status']=='pass'
 assert d.read(p/'qualification.json')['status']=='commands_pass'
 d.check('baseline')
 plan=d.read(p/'measurement-plan.json')
 assert len(plan['rows'])==72 and sum(r['samples'] for r in plan['rows'])==2160
 d.write(p/'capture-started.json',dict(plan=d.desc(p/'measurement-plan.json'),admission=d.desc(p/'admission.json'),driver=d.desc(p/'capture.py'),freeze=d.desc(p/'freeze-baseline.json')))
 rows=[]
 for i,row in enumerate(plan['rows']):
  result=d.workload('baseline',f'native-{i:03}',row['case'],row['cache_state'],row['samples'],row['warmup'])
  rows.append(dict(plan=row,result=result))
  if result['exit_code']!=0:
   d.write(p/'capture-failed.json',dict(rows=rows,failed_index=i));raise RuntimeError(f'formal capture failed at {i}; no retry')
 d.write(p/'capture.json',dict(status='commands_pass',rows=rows,report_count=len(rows),sample_count=2160))
 print('formal commands PASS: 72 reports, 2160 samples; independent analysis pending',flush=True)

if __name__=='__main__':main()
