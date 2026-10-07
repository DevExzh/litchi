#!/usr/bin/env python3
"""Freeze a reviewed, bound prospective plan for the coordinator's capture."""
import datetime, hashlib, json
from pathlib import Path
ROOT=Path('/home/zhuhe/litchi-performance-0961')
def desc(p):
 b=p.read_bytes();return {'path':str(p),'bytes':len(b),'mode':p.stat().st_mode&0o777,'sha256':hashlib.sha256(b).hexdigest()}
def write(p,v):
 with p.open('x') as f:json.dump(v,f,indent=2);f.write('\n')
plan_path=ROOT/'protocol/plan-bound-prospective-v2.json'
plan=json.loads(plan_path.read_text())
result=json.loads((ROOT/'receipts/candidate-managed-v1-build-result.json').read_text())
assert result['returncode']==0 and result['source_unchanged'] and result['retained_executables']
assert json.loads((ROOT/'protocol/preflight-v2.json').read_text())['status']=='check-plan-pass'
assert not json.loads((ROOT/'protocol/descriptor-preflight-v2.json').read_text())['mismatches']
for lane in ('native','allocator'):assert not Path(plan['capture_root'][lane]).exists()
receipt={'schema':'litchi.performance.0961.root-go.v1','status':'GO','created_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'scope':'One screening experiment only; sequential native then allocator; no retries; no adoption','bound_prospective_plan':desc(plan_path),'preflight':desc(ROOT/'protocol/preflight-v2.json'),'descriptor_preflight':desc(ROOT/'protocol/descriptor-preflight-v2.json'),'build_result':desc(ROOT/'receipts/candidate-managed-v1-build-result.json'),'source_review':'Independent staging_review: implementation and quality accepted for screening; exact8/256 and ordinary/fallback controls remain adoption gates'}
go=ROOT/'protocol/root-go-v1.json';write(go,receipt)
plan['status']='authorized';plan['root_go']={'status':'GO','receipt':desc(go)}
plan['status_note']='Root authorized reviewed 0961 screening; no measurements started at freeze.'
write(ROOT/'protocol/plan-final-v1.json',plan)
print(json.dumps({'status':'GO','plan':desc(ROOT/'protocol/plan-final-v1.json')}))
