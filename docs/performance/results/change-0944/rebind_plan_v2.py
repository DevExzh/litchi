#!/usr/bin/env python3
from pathlib import Path
import hashlib,json,os
PACKET=Path('/home/zhuhe/litchi-performance-0944')
PLAN=PACKET/'docs/performance/results/change-0944/plan.json'
def digest(path):
 h=hashlib.sha256(); n=0
 with path.open('rb') as f:
  for b in iter(lambda:f.read(1<<20),b''): h.update(b); n+=len(b)
 return n,h.hexdigest()
def rel(path): return str(path.resolve().relative_to(PACKET.resolve()))
def d(path):
 n,h=digest(path);return {'path':rel(path),'bytes':n,'sha256':h}
plan=json.loads(PLAN.read_text())
for key,path in {
 'build_process_group_audit':PACKET/'build/build-process-group-audit.json',
 'source_checker':PACKET/'tools/source_check_v2.py',
 'semantic_validator':PACKET/'tools/semantic_v2.py',
 'reader':PACKET/'docs/performance/results/change-0845/reader_v2.py',
 'raw_analyzer':PACKET/'tools/raw_analyzer_v2.py',
 'capture_runner':PACKET/'tools/capture_v2.py',
 'retained_helpers':PACKET/'docs/performance/results/change-0944/retained-helper-custody.json',
 'native_context':PACKET/'references/native-context-0900.json',
 'prior_semantic_validator':PACKET/'docs/performance/results/change-0944/helpers/retained-0705-semantic-analyze.py',
 'raw_adapter':PACKET/'docs/performance/results/change-0944/analyze_v2_adapter.py',
 'build_runner':PACKET/'tools/build_native.py',
 'audit_controls':PACKET/'tools/audit_controls_v2.py',
 'audit_control_results':PACKET/'controls/audit-controls-v2.json',
}.items(): plan[key]=d(path)
plan['plan_generator']=d(PACKET/'tools/prepare_plan.py')
# Atomic rewrite is before any workload and deliberately has no self hash.
tmp=PLAN.with_suffix('.tmp'); tmp.write_text(json.dumps(plan,indent=2,sort_keys=True)+'\n'); os.replace(tmp,PLAN)
print(json.dumps({'plan':rel(PLAN),'sha256':hashlib.sha256(PLAN.read_bytes()).hexdigest(),'script_hashes':{k:v['sha256'] for k,v in plan.items() if k in ('source_checker','semantic_validator','raw_analyzer','capture_runner')}}))
