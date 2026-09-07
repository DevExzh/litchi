#!/usr/bin/env python3
"""Replay an exported bundle and reject semantic/custody corruption without its outer seal."""
import argparse,json,shutil,subprocess,sys,tempfile
from pathlib import Path
ROOT=Path(__file__).resolve().parent
def run(root,args=()):
    r=subprocess.run([sys.executable,'-B',str(root/'verify.py'),*args],cwd=root,capture_output=True,text=True)
    return {'exit_code':r.returncode,'stdout':r.stdout,'stderr':r.stderr}
p=argparse.ArgumentParser();p.add_argument('--stage',choices=['precleanup','aftercleanup'],required=True);a=p.parse_args()
args=['--sealed']+(['--cleanup'] if a.stage=='aftercleanup' else [])
rows=[]
with tempfile.TemporaryDirectory(prefix='litchi-goal-0453-portable-') as scratch:
    root=Path(scratch)/'export';shutil.copytree(ROOT,root)
    valid=run(root,args);assert valid['exit_code']==0,valid
    cases=[('measurements.json',lambda v:v['rows'][0].__setitem__('api_sum_p50_ms',0)),
           ('checks/final-opc-r2.json',lambda v:v.__setitem__('source_unchanged',False)),
           ('checks/final-opc-r2.json',lambda v:v.__setitem__('passed_tests',472)),
           ('checks/fuzz-smoke.json',lambda v:v['argv'].__setitem__(3,'-seed=1')),
           ('decision.json',lambda v:v.__setitem__('native_coverage_promoted',True)),
           ('candidate-build.json',lambda v:v['binaries']['normal'].__setitem__('sha256','0'*64)),
           ('runs/0/receipt.json',lambda v:v['binary'].__setitem__('sha256','0'*64)),
           ('runs/0/report.json',lambda v:v['samples_raw'][0].__setitem__('exact_output_verified',False)),
           ('regression-review.json',lambda v:v.__setitem__('flags',[])),
           ('checks/fuzz-prepared.json',lambda v:v['inputs'][0].__setitem__('sha256','0'*64)),
           ('checks/fuzz-artifacts.json',lambda v:v.__setitem__('files',0))]
    cases += [('confirmation-summary.json',lambda v:v.__setitem__('samples',0)),('runs/16/report.json',lambda v:v['samples_raw'][0]['timings']['plan_allocation_metrics'].__setitem__('allocated_bytes',0)),('runs/16/report.json',lambda v:v.__setitem__('instrumentation','none')),('protocol.json',lambda v:v.__setitem__('samples',1))]
    if a.stage=='aftercleanup':cases += [('checks/fuzz-cleanup.json',lambda v:v.__setitem__('status','failed')),('checks/binary-cleanup.json',lambda v:v.__setitem__('temporary_directory_absent',False))]
    inner=['--cleanup'] if a.stage=='aftercleanup' else []
    for n,mutate in cases:
        path=root/n;raw=path.read_bytes();v=json.loads(raw);mutate(v);path.write_text(json.dumps(v))
        result=run(root,inner);assert result['exit_code']!=0,(n,result);rows.append({'path':n,**result});path.write_bytes(raw)
    path=root/'candidate/after-source_cross_copy.rs.txt';raw=path.read_bytes();path.write_bytes(raw+b'\n// altered candidate\n')
    result=run(root,inner);assert result['exit_code']!=0;rows.append({'path':str(path.relative_to(root)),**result});path.write_bytes(raw)
    (root/'unsealed.txt').write_text('unsealed member')
    result=run(root,args);assert result['exit_code']!=0;rows.append({'path':'unsealed.txt',**result})
receipt={'change':453,'stage':a.stage,'status':'pass','valid_export':valid,'rejected':rows,'scratch_removed':not Path(scratch).exists()}
with (ROOT/'checks'/('probes-'+a.stage+'.json')).open('x') as f:f.write(json.dumps(receipt,indent=2)+'\n')
print(json.dumps({'status':'pass','stage':a.stage,'rejected':len(rows),'scratch_removed':True}))
