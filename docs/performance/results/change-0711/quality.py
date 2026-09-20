#!/usr/bin/env python3
import importlib.util,json,subprocess,time
from pathlib import Path
P=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('c0711',P/'custody.py');C=importlib.util.module_from_spec(spec);spec.loader.exec_module(C)
def main():
    assert not (P/'quality.json').exists()
    source=C.census();(P/'quality-source.json').write_text(json.dumps(source,indent=2)+'\n')
    common=['--locked','--target-dir',str(C.TARGET),'-j','2']
    jobs=[('fmt',['cargo','fmt','--all','--check']),('docx-tests',['cargo','test','-p','litchi-docx',*common,'--all-features','--all-targets']),('docx-clippy',['cargo','clippy','-p','litchi-docx',*common,'--all-features','--all-targets','--','-D','warnings']),('docx-doctests',['cargo','test','-p','litchi-docx',*common,'--all-features','--doc']),('docx-rustdoc',['env','RUSTDOCFLAGS=-D warnings','cargo','doc','-p','litchi-docx',*common,'--all-features','--no-deps'])]
    rows=[]
    for name,cmd in jobs:
        log=P/('quality-'+name+'.log');start=time.monotonic()
        with log.open('w') as f:r=subprocess.run(cmd,cwd=C.ROOT,stdout=f,stderr=subprocess.STDOUT)
        assert C.census()==source
        rows.append(dict(name=name,command=cmd,exit_code=r.returncode,seconds=time.monotonic()-start,log=log.name,log_sha256=C.sha(log),source_manifest_sha256=C.sha(P/'quality-source.json')))
        (P/'quality.json').write_text(json.dumps(rows,indent=2)+'\n');print(name,r.returncode,flush=True);assert r.returncode==0
if __name__=='__main__':main()
