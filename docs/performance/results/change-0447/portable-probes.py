#!/usr/bin/env python3
"""Verify an exported bundle and reject custody mutations without workload execution."""
import argparse,json,shutil,subprocess,sys,tempfile
from pathlib import Path
ROOT=Path(__file__).resolve().parent

def run(directory,extra=()):
    p=subprocess.run([sys.executable,'-B',str(directory/'verify.py'),*extra],cwd=directory,capture_output=True,text=True)
    return {'exit_code':p.returncode,'stdout':p.stdout,'stderr':p.stderr}
def main():
    parser=argparse.ArgumentParser();parser.add_argument('--stage',choices=['precleanup','aftercleanup'],required=True);a=parser.parse_args()
    out=ROOT/'checks'/('probes-'+a.stage+'.json');assert not out.exists()
    extra=['--sealed']+(['--cleanup'] if a.stage=='aftercleanup' else [])
    rows=[]
    with tempfile.TemporaryDirectory(prefix='litchi-goal-0447-portable-probes-') as scratch:
        directory=Path(scratch)/'export';shutil.copytree(ROOT,directory)
        valid=run(directory,extra);assert valid['exit_code']==0,valid
        cases=[('protocol.json',lambda v:v.__setitem__('samples',29)),
               ('build.json',lambda v:v['binary'].__setitem__('sha256','0'*64)),
               ('runs/0/receipt.json',lambda v:v.__setitem__('source_unchanged',False)),
               ('runs/0/receipt.json',lambda v:v['argv'].__setitem__(2,'3'))]
        if a.stage=='aftercleanup':cases.append(('checks/cleanup-inventory.json',lambda v:v.__setitem__('status','failed')))
        for number,(name,mutate) in enumerate(cases):
            path=directory/name;original=path.read_bytes();value=json.loads(original);mutate(value);path.write_text(json.dumps(value))
            result=run(directory,extra);assert result['exit_code']!=0,(name,result)
            rows.append({'mutation':number,'path':name,**result});path.write_bytes(original)
        unexpected=directory/'unexpected.txt';unexpected.write_text('unsealed member')
        result=run(directory,extra);assert result['exit_code']!=0,result
        rows.append({'mutation':'extra-seal-member',**result})
    value={'change':447,'status':'pass','stage':a.stage,'exported_bundle_verification':valid,'rejected':rows,'scope':'portable verification runs no Git, original binaries, workload or profiler; scratch copy removed before receipt publication'}
    with out.open('x') as stream:stream.write(json.dumps(value,indent=2)+'\n')
    print(json.dumps({'status':'pass','rejected':len(rows),'stage':a.stage}))
if __name__=='__main__':main()
