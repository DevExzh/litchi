#!/usr/bin/env python3
"""Verify retained source/probe/fixture/raw bindings and compare allocation/I/O."""
import hashlib,json,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
def read(p):return json.loads(p.read_text())
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def untimed(v):
 if isinstance(v,dict):return {k:untimed(x) for k,x in v.items() if 'elapsed' not in k}
 if isinstance(v,list):return [untimed(x) for x in v]
 return v
base=read(P/'baseline.json')
for n,h in base['constraints_sha256'].items():assert sha(ROOT/n)==h,n
for phase in ['baseline','candidate']:
 d=P/'costs'/phase;m=read(d/'manifest.json')
 for n,h in m['raw_sha256'].items():assert sha(d/n)==h,n
 for field in ['probe_sha256','corpus']:
  for n,h in m[field].items():assert sha(ROOT/n)==h,n
 assert sha(P/'cases.json')==m['cases_sha256']
 for n,h in m['source_sha256'].items():
  actual=sha(ROOT/n) if phase=='candidate' else hashlib.sha256(subprocess.check_output(['git','show',base['baseline_head']+':'+n],cwd=ROOT)).hexdigest()
  assert actual==h,n
 for n,h in m['binary_sha256'].items():
  f=Path('/home/zhuhe/code/litchi-target-0690-'+('before' if phase=='baseline' else 'after'))/'release'/n
  if f.exists():assert sha(f)==h,n
 assert all(r['exit_code']==0 for r in read(d/'commands.json'))
rows=[]
io_rows=[]
for c in read(P/'cases.json'):
 name=f"counts-{c['case']}.json"
 before=read(P/'costs/baseline'/name);after=read(P/'costs/candidate'/name)
 assert before['open_metrics']==after['open_metrics'] and before['open_error']==after['open_error']
 assert len(before['queries'])==len(after['queries'])==8
 for i,(b,a) in enumerate(zip(before['queries'],after['queries'])):
  assert b['outcome']==a['outcome'],(name,i)
  assert b['metrics']==a['metrics'],(name,i)
 io_rows.append(dict(case=c['case'],baseline=untimed(before),candidate=untimed(after)))
 for mode in ['owned','file']:
  for op in ['q1','q2','q3','q8']:
   phases={}
   for phase in ['baseline','candidate']:
    values=[read(P/'costs'/phase/f"alloc-{c['case']}-{mode}-{op}-{n}.json") for n in range(3)]
    assert all(v==values[0] for v in values), (phase,c['case'],mode,op)
    phases[phase]=values[0]
   gauges={'allocation_calls','allocated_bytes','deallocation_calls','deallocated_bytes','peak_live_delta','retained_live_delta'}
   assert {k:v for k,v in phases['baseline'].items() if k not in gauges}=={k:v for k,v in phases['candidate'].items() if k not in gauges},(c['case'],mode,op)
   assert phases['baseline']==phases['candidate'],(c['case'],mode,op)
   rows.append(dict(case=c['case'],mode=mode,operation=op,**phases))
(P/'io-comparison.json').write_text(json.dumps(io_rows,indent=2)+'\n')
(P/'allocation-comparison.json').write_text(json.dumps(rows,indent=2)+'\n')
print('PASS: source/probe/raw/constraint bindings, exact all-query outcome/I/O, exact allocation parity in all build and non-build groups,',len(rows),'allocation groups with three identical repeats per phase')
for r in rows:
 if r['operation']=='q8' and r['mode']=='owned':
  print(r['case'],[(k,r['baseline'][k],r['candidate'][k]) for k in ['allocation_calls','allocated_bytes','peak_live_delta','retained_live_delta']])
