#!/usr/bin/env python3
import argparse,json,shutil,subprocess,sys,tempfile
from pathlib import Path
ROOT=Path(__file__).resolve().parent
def run(root,args=()):
    r=subprocess.run([sys.executable,'-B',str(root/'verify.py'),*args],cwd=root,capture_output=True,text=True)
    return {'exit_code':r.returncode,'stdout':r.stdout,'stderr':r.stderr}
p=argparse.ArgumentParser();p.add_argument('--stage',choices=['precleanup','aftercleanup'],required=True);a=p.parse_args()
args=['--sealed']+(['--cleanup'] if a.stage=='aftercleanup' else [])
rows=[]
with tempfile.TemporaryDirectory(prefix='litchi-goal-0451-portable-') as scratch:
    root=Path(scratch)/'export';shutil.copytree(ROOT,root)
    valid=run(root,args);assert valid['exit_code']==0,valid
    cases=[('measurements.json',lambda v:v['rows'][0].__setitem__('combined_bytes',0)),
           ('checks/final-opc-tests.json',lambda v:v.__setitem__('source_unchanged',False)),
           ('checks/final-opc-tests.json',lambda v:v.__setitem__('passed_tests',467)),
           ('checks/fuzz-smoke.json',lambda v:v['argv'].__setitem__(3,'-seed=1')),
           ('decision.json',lambda v:v.__setitem__('pptx_adopted',True)),
           ('checks/fuzz-prepared.json',lambda v:v['inputs'][0].__setitem__('sha256','0'*64)),
           ('checks/fuzz-artifacts.json',lambda v:v.__setitem__('files',0))]
    if a.stage=='aftercleanup':cases.append(('checks/fuzz-cleanup.json',lambda v:v.__setitem__('status','failed')))
    # Run semantic/custody mutations without the outer seal so each relevant
    # verifier check, rather than only a file checksum, must reject them.
    inner=['--cleanup'] if a.stage=='aftercleanup' else []
    for n,mutate in cases:
        path=root/n;raw=path.read_bytes();v=json.loads(raw);mutate(v);path.write_text(json.dumps(v))
        result=run(root,inner);assert result['exit_code']!=0,(n,result);rows.append({'path':n,**result});path.write_bytes(raw)
    path=root/'candidate/after-source_backed.rs.txt';raw=path.read_bytes();path.write_bytes(raw+b'\n// changed tested source\n')
    result=run(root,inner);assert result['exit_code']!=0;rows.append({'path':str(path.relative_to(root)),**result});path.write_bytes(raw)
    (root/'unsealed.txt').write_text('unsealed member')
    result=run(root,args);assert result['exit_code']!=0;rows.append({'path':'unsealed.txt',**result})
receipt={'change':451,'stage':a.stage,'status':'pass','valid_export':valid,'rejected':rows,'scratch_removed':not Path(scratch).exists()}
with (ROOT/'checks'/('probes-'+a.stage+'.json')).open('x') as f:f.write(json.dumps(receipt,indent=2)+'\n')
print(json.dumps({'status':'pass','stage':a.stage,'rejected':len(rows),'scratch_removed':True}))
