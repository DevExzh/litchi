#!/usr/bin/env python3
"""Serialized source-bound checks for the custom-properties preservation fix."""
import importlib.util
import json
from pathlib import Path
import subprocess
import sys
import time
P=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('custody0710',P/'custody.py');C=importlib.util.module_from_spec(spec);spec.loader.exec_module(C)
def main():
    stage=sys.argv[1];assert stage in ['baseline','candidate']
    out=P/stage;out.mkdir(exist_ok=True);assert not (out/'checks.json').exists()
    source=C.census();(out/'source.json').write_text(json.dumps(source,indent=2)+'\n')
    common=['--locked','--target-dir',str(C.TARGET),'-j','2']
    if stage=='baseline':
        jobs=[('custom-properties',['cargo','test','-p','litchi-docx',*common,'--all-features','--test','custom_properties'])]
    else:
        jobs=[('fmt',['cargo','fmt','--all','--check']),('docx-all-features',['cargo','test','-p','litchi-docx',*common,'--all-features','--all-targets']),('docx-clippy',['cargo','clippy','-p','litchi-docx',*common,'--all-features','--all-targets','--','-D','warnings'])]
    rows=[]
    for name,cmd in jobs:
        log=out/(name+'.log');start=time.monotonic()
        with log.open('w') as f:r=subprocess.run(cmd,cwd=C.ROOT,stdout=f,stderr=subprocess.STDOUT)
        assert C.census()==source
        rows.append(dict(name=name,command=cmd,exit_code=r.returncode,seconds=time.monotonic()-start,log_sha256=C.sha(log),source_sha256=C.sha(out/'source.json')))
        (out/'checks.json').write_text(json.dumps(rows,indent=2)+'\n');print(stage,name,r.returncode,flush=True)
        if stage=='candidate':assert r.returncode==0
    if stage=='baseline':assert rows[0]['exit_code']==101
if __name__=='__main__':main()
