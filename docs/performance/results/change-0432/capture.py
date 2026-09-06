#!/usr/bin/env python3
"""Execute the frozen twelve fresh-process XLSX streaming matrix, serially."""
import datetime
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import subprocess
import sys
ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()
def now():
    return datetime.datetime.now(datetime.timezone.utc).isoformat()
def main():
    protocol = json.loads((ROOT/'protocol.json').read_text())
    build = json.loads((ROOT/'build.json').read_text())
    spec = importlib.util.spec_from_file_location('custody', ROOT/'check.py')
    custody = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(custody)
    assert custody.sources() == build['source_manifest']
    assert protocol['cpu'] in os.sched_getaffinity(0)
    assert sha(ROOT/'protocol.json') == build['protocol_sha256']
    assert sha(ROOT/'verify-report.py') == build['verifier_sha256']
    subprocess.run(['git','diff','--exit-code','HEAD','--'],cwd=REPO,check=True,stdout=subprocess.DEVNULL)
    status_before=subprocess.check_output(['git','status','--short'],cwd=REPO,text=True).splitlines()
    assert status_before==['?? docs/GOAL.md','?? docs/performance/results/change-0432/']
    directory = ROOT/'captures'; directory.mkdir(exist_ok=False)
    index = []
    for lane in protocol['order']:
        name = '-'.join([lane['mode'],lane['shape'],lane['repeat'].lower()])
        binary = build['binaries'][lane['mode']]
        assert sha(Path(binary['path'])) == binary['sha256']
        report = directory/(name+'.json'); catalog = directory/(name+'-catalog.json')
        log = directory/(name+'.log'); resource = directory/(name+'-resource.log')
        receipt = directory/(name+'-receipt.json')
        argv = ['taskset','-c',str(protocol['cpu']),'/usr/bin/time','-v','-o',str(resource),binary['path'],'--case',protocol['selector'],'--semantic-shape',lane['shape'],'--workers','1','--samples',str(protocol['samples']),'--warmup',str(protocol['warmups']),'--json',str(report),'--corpus-manifest',str(catalog)]
        row={'change':432,'name':name,'lane':lane,'argv':argv,'revision':build['revision'],'source_manifest':build['source_manifest'],'binary':binary,'protocol_sha256':sha(ROOT/'protocol.json'),'driver_sha256':sha(Path(__file__)),'verifier_sha256':sha(ROOT/'verify-report.py'),'started_utc':now(),'status':'running'}
        receipt.write_text(json.dumps(row,indent=2)+'\n'); print('START '+name,flush=True)
        try:
            with log.open('xb') as stream:
                result=subprocess.run(argv,cwd=REPO,stdout=stream,stderr=subprocess.STDOUT)
            row['exit_code']=result.returncode
            assert result.returncode == 0
            subprocess.run([sys.executable,'-B',str(ROOT/'verify-report.py'),'--report',str(report),'--mode',lane['mode'],'--shape',lane['shape']],check=True)
            row['status']='pass'
        finally:
            if row['status']=='running': row['status']='failed'
            row['finished_utc']=now()
            row['artifacts']={str(p.relative_to(ROOT)):{'sha256':sha(p),'bytes':p.stat().st_size} for p in [report,catalog,log,resource] if p.is_file()}
            receipt.write_text(json.dumps(row,indent=2)+'\n')
        index.append(str(receipt.relative_to(ROOT))); print('FINISH '+name,flush=True)
    assert len(index)==12
    assert custody.sources()==build['source_manifest']
    subprocess.run(['git','diff','--exit-code','HEAD','--'],cwd=REPO,check=True,stdout=subprocess.DEVNULL)
    status_after=subprocess.check_output(['git','status','--short'],cwd=REPO,text=True).splitlines()
    assert status_after==status_before
    (ROOT/'capture-state.json').write_text(json.dumps({'status':'pass','tracked_tree_clean_before_and_after':True,'status_before':status_before,'status_after':status_after,'scope':'Harness git_worktree_dirty=true includes untracked user goal and this evidence bundle; committed sources are clean and full code manifests stay equal.'},indent=2)+'\n')
    (ROOT/'capture-index.json').write_text(json.dumps(index,indent=2)+'\n')
if __name__=='__main__': main()
