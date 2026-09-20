#!/usr/bin/env python3
"""Serialize bounded quality checks with exact source custody."""
import importlib.util
import json
from pathlib import Path
import subprocess
import sys
import time
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
spec=importlib.util.spec_from_file_location('build_custody',P/'build.py')
B=importlib.util.module_from_spec(spec);spec.loader.exec_module(B)

def main():
    phase=sys.argv[1]
    common=['--locked','--target-dir',str(B.TARGET),'-j','2']
    focused=[
      ('validation',['cargo','test','-p','litchi-xlsx',*common,'--lib','cell_values::validation']),
      ('facts',['cargo','test','-p','litchi-xlsx',*common,'--lib','cell_values::facts_oracle_tests']),
      ('planning-errors',['cargo','test','-p','litchi-xlsx',*common,'--test','source_backed_cell_values','planning_error_order'])]
    full=[
      ('fmt',['cargo','fmt','--all','--check']),
      ('xlsx-tests',['cargo','test','-p','litchi-xlsx',*common,'--all-features','--all-targets']),
      ('xlsx-check',['cargo','check','-p','litchi-xlsx',*common,'--all-features','--all-targets']),
      ('xlsx-clippy',['cargo','clippy','-p','litchi-xlsx',*common,'--all-features','--all-targets','--','-D','warnings']),
      ('facade-tests',['cargo','test','-p','litchi',*common,'--no-default-features','--features','xlsx,encryption,vba-inspection','--lib','--tests'])]
    assert phase in ['baseline-focused','candidate-focused','candidate-full','restored-focused','restored-quality']
    checks=full if phase=='candidate-full' else ([full[0],full[2],full[3]] if phase=='restored-quality' else focused)
    output=P/('quality-'+phase+'.json');assert not output.exists()
    source=B.census();source_path=P/('quality-'+phase+'.source.json')
    source_path.write_text(json.dumps(source,indent=2)+'\n');rows=[]
    for name,command in checks:
      log=P/(phase+'-'+name+'.log');assert not log.exists()
      start=time.monotonic()
      with log.open('w') as stream:r=subprocess.run(command,cwd=ROOT,stdout=stream,stderr=subprocess.STDOUT)
      assert B.census()==source,'source changed during quality check'
      rows.append(dict(name=name,command=command,exit_code=r.returncode,seconds=time.monotonic()-start,log=log.name,log_sha256=B.sha(log),source_manifest=source_path.name,source_manifest_sha256=B.sha(source_path)))
      output.write_text(json.dumps(rows,indent=2)+'\n');print(phase,name,r.returncode,flush=True)
      assert r.returncode==0
if __name__=='__main__':main()
