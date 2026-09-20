#!/usr/bin/env python3
"""Build and run a separately source-bound public DOCX semantic oracle."""
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import shutil
import subprocess
import time
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
O=P/'oracle'
spec=importlib.util.spec_from_file_location('custody0710',P/'custody.py')
B=importlib.util.module_from_spec(spec);spec.loader.exec_module(B)
def main():
    assert not (O/'result.json').exists()
    source=B.census()
    probe={str(f.relative_to(ROOT)):B.sha(f) for f in [O/'Cargo.toml',O/'Cargo.lock',O/'src/main.rs']}
    (O/'source.json').write_text(json.dumps(probe,indent=2)+'\n')
    (O/'workspace-source.json').write_text(json.dumps(source,indent=2)+'\n')
    common=['--manifest-path',str(O/'Cargo.toml'),'--locked','--target-dir',str(B.TARGET),'-j','2']
    commands=[('fmt',['cargo','fmt','--manifest-path',str(O/'Cargo.toml'),'--','--check']),('build',['cargo','build','--release',*common]),('clippy',['cargo','clippy',*common,'--','-D','warnings'])]
    rows=[]
    for name,command in commands:
        log=O/(name+'.log');assert not log.exists()
        start=time.monotonic()
        with log.open('w') as f:r=subprocess.run(command,cwd=ROOT,stdout=f,stderr=subprocess.STDOUT)
        assert B.census()==source
        assert all(B.sha(ROOT/n)==h for n,h in probe.items())
        rows.append(dict(name=name,command=command,exit_code=r.returncode,seconds=time.monotonic()-start,log=log.name,log_sha256=B.sha(log)))
        (O/'checks.json').write_text(json.dumps(rows,indent=2)+'\n');print('oracle',name,r.returncode,flush=True)
        assert r.returncode==0
    B.BIN.mkdir(exist_ok=True)
    binary=B.BIN/'oracle';shutil.copy2(B.TARGET/'release/docx-ordinary-save-oracle-0709',binary)
    tmp=ROOT.parent/'litchi-0710-fs'/'oracle';tmp.mkdir(parents=True,exist_ok=True)
    env=dict(os.environ,TMPDIR=str(tmp),LITCHI_GIT_REV=json.loads((P/'revision.json').read_text())['revision'])
    command=[str(binary),'--repo-root',str(ROOT),'--marker','litchi-perf-0638-ordinary-save','--output',str(O/'report.json')]
    with (O/'run.stdout').open('w') as out,(O/'run.stderr').open('w') as err:r=subprocess.run(command,cwd=ROOT,env=env,stdout=out,stderr=err)
    assert B.census()==source and all(B.sha(ROOT/n)==h for n,h in probe.items())
    result=dict(command=command,exit_code=r.returncode,binary=dict(path=str(binary),sha256=B.sha(binary),bytes=binary.stat().st_size),probe_manifest_sha256=B.sha(O/'source.json'),workspace_manifest_sha256=B.sha(O/'workspace-source.json'),artifacts={str(f.relative_to(O)):B.sha(f) for f in [O/'run.stdout',O/'run.stderr',O/'report.json',*sorted((O/'artifacts').glob('*.docx'))] if f.exists()},expected_preservation_failure=False,temporary_root=str(tmp))
    (O/'result.json').write_text(json.dumps(result,indent=2)+'\n');print('oracle semantic',r.returncode,flush=True)
    report=json.loads((O/'report.json').read_text())
    assert r.returncode==0 and report['oracle_pass'] is True
    assert report['successful_fixture']['fixture_pass'] is True
    assert report['refusal_fixture']['fixture_pass'] is True
    assert all(route['checks']['no_non_main_decoded_payload_changed'] for route in report['refusal_fixture']['no_edit_routes'])
    print('Strict preservation oracle and no-edit controls PASS',flush=True)
if __name__=='__main__':main()
