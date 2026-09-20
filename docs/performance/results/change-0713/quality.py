#!/usr/bin/env python3
"""Run shared MCE and dependent OOXML quality checks serially."""
import importlib.util,json,subprocess,time,sys
from pathlib import Path
P=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('custody0713quality',P/'custody.py');C=importlib.util.module_from_spec(spec);spec.loader.exec_module(C)
def main():
    assert not (P/'quality.json').exists()
    source=C.census();(P/'quality-source.json').write_text(json.dumps(source,indent=2)+'\n')
    common=['--locked','--target-dir',str(C.TARGET),'-j','2']
    names=['litchi-ooxml-common'] if '--focused' in sys.argv else ['litchi-ooxml-common','litchi-drawingml','litchi-spreadsheet-drawing','litchi-docx','litchi-xlsx','litchi-pptx','litchi-xlsb']
    packages=[arg for name in names for arg in ['-p',name]]
    jobs=[('fmt',['cargo','fmt','--all','--check']),('tests',['cargo','test',*packages,*common,'--all-features','--all-targets']),('clippy',['cargo','clippy',*packages,*common,'--all-features','--all-targets','--','-D','warnings']),('doctests',['cargo','test',*packages,*common,'--all-features','--doc']),('rustdoc',['env','RUSTDOCFLAGS=-D warnings','cargo','doc',*packages,*common,'--all-features','--no-deps'])]
    rows=[]
    for name,cmd in jobs:
        log=P/('quality-'+name+'.log');start=time.monotonic()
        with log.open('w') as f:r=subprocess.run(cmd,cwd=C.ROOT,stdout=f,stderr=subprocess.STDOUT)
        assert C.census()==source
        rows.append(dict(name=name,command=cmd,exit_code=r.returncode,seconds=time.monotonic()-start,log=log.name,log_sha256=C.sha(log),source_manifest_sha256=C.sha(P/'quality-source.json')))
        (P/'quality.json').write_text(json.dumps(rows,indent=2)+'\n');print(name,r.returncode,flush=True);assert r.returncode==0
if __name__=='__main__':main()
