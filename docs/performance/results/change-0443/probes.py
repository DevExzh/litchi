#!/usr/bin/env python3
"""Accept a portable copied control and reject separately corrupted copies."""
import argparse
import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile

ROOT = Path(__file__).resolve().parent


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def load(path):
    return json.loads(path.read_text())


def write(path, value):
    path.write_text(json.dumps(value,indent=2)+'\n')


def refresh(root):
    (root/'SHA256SUMS').write_text(''.join(f'{sha(path)}  {path.relative_to(root)}\n' for path in sorted(root.rglob('*')) if path.is_file() and path.name!='SHA256SUMS'))


def run(root, cleanup):
    command=[sys.executable,'-B',str(root/'verify.py'),'--sealed']+(['--cleanup'] if cleanup else [])
    result=subprocess.run(command,capture_output=True,text=True,cwd=root.parent)
    return {'exit_code':result.returncode,'output':(result.stdout+result.stderr)[-2000:]}


def mutate_json(root, name, mutation):
    path=root/name;value=load(path);mutation(value);write(path,value)


def mutate_report(root, allocation=False):
    receipt=root/('runs/A1/formal/A1-before-allocator-large-r1-receipt.json' if allocation else 'runs/A1/formal/A1-before-normal-tiny-r1-receipt.json')
    row=load(receipt);path=root/row['artifacts']['report']['path'];report=load(path)
    if allocation:
        report['results'][0]['operation_metrics']['allocation']['live_bytes_after']['values'][0]+=1
    else:
        report['results'][0]['source']['odp_append']['append_exactly_one_verified']=False
    write(path,report)
    row['artifacts']['report'].update(bytes=path.stat().st_size,sha256=sha(path))
    write(receipt,row)


def main():
    parser=argparse.ArgumentParser();parser.add_argument('--tag',required=True);parser.add_argument('--cleanup',action='store_true');args=parser.parse_args()
    assert args.tag.replace('-','').isalnum()
    output=ROOT/'checks'/f'probes-{args.tag}.json';assert not output.exists()
    before=sha(ROOT/'SHA256SUMS')
    cases=[
        ('oracle_verifier',lambda r:(r/'oracle/verify-report.py').write_bytes((r/'oracle/verify-report.py').read_bytes()+b'\n# changed\n')),
        ('protocol_cpu',lambda r:mutate_json(r,'protocol.json',lambda x:x.update(cpu=3))),
        ('build_binary',lambda r:mutate_json(r,'before/build.json',lambda x:x['binaries']['normal'].update(sha256='0'*64))),
        ('phase_order',lambda r:write(r/'runs/A1/formal/capture-index.json',list(reversed(load(r/'runs/A1/formal/capture-index.json'))))),
        ('capture_source',lambda r:mutate_json(r,'runs/A1/formal/A1-before-normal-tiny-r1-receipt.json',lambda x:x['source_after'].update(sha256='0'*64))),
        ('report_semantics',lambda r:mutate_report(r)),
        ('allocator_balance',lambda r:mutate_report(r,True)),
        ('profile_argv',lambda r:mutate_json(r,'profiles/after/stat/formal/receipt.json',lambda x:x['argv'].__setitem__(2,'3'))),
        ('candidate_source',lambda r:(r/'candidate/after-scanner.rs.txt').write_bytes((r/'candidate/after-scanner.rs.txt').read_bytes()+b'\n// altered\n')),
        ('gate_disclosure',lambda r:mutate_json(r,'decision.json',lambda x:x.update(practical_gate_met=not x['practical_gate_met']))),
        ('required_check',lambda r:(r/'checks/candidate-odp-tests.json').unlink()),
    ]
    if args.cleanup:
        cases.append(('cleanup_goal',lambda r:mutate_json(r,'checks/cleanup-inventory.json',lambda x:x.update(goal_sha256='0'*64))))
    results=[]
    with tempfile.TemporaryDirectory(prefix='litchi-0443-probes-') as temporary:
        control=Path(temporary)/'control';shutil.copytree(ROOT,control)
        positive=run(control,args.cleanup);assert positive['exit_code']==0,positive
        for name,mutation in cases:
            copy=Path(temporary)/name;shutil.copytree(ROOT,copy);mutation(copy);refresh(copy)
            result=run(copy,args.cleanup);assert result['exit_code']!=0,name
            results.append({'name':name,'status':'rejected',**result})
    assert sha(ROOT/'SHA256SUMS')==before
    write(output,{'status':'pass','copied_control':positive,'probes':results,'count':len(results),'inventory_unchanged':True,'inventory_sha256':before,'verifier_sha256':sha(ROOT/'verify.py'),'scope':'independent copied bundles; mutated report artifact hashes and whole inventories refreshed before verification; source bundle untouched until this completed proof is written'})
    print(json.dumps({'status':'pass','count':len(results)}))


if __name__=='__main__':
    main()
