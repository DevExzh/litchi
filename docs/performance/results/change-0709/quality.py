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
    common=['--manifest-path','tools/perf-baseline/Cargo.toml','--locked','--target-dir',str(B.TARGET),'-j','2']
    checks=[
      ('fmt',['cargo','fmt','--manifest-path','tools/perf-baseline/Cargo.toml','--all','--check']),
      ('ordinary-save-tests',['cargo','test',*common,'--release','--lib','ordinary_save::tests']),
      ('harness-clippy',['cargo','clippy',*common,'--all-features','--all-targets','--','-D','warnings']),
      ('source-policy',['python3','tools/test_perf_baseline_source_policy.py'])]
    output=P/'quality.json';assert not output.exists()
    source=B.census();source_path=P/'quality-source.json'
    source_path.write_text(json.dumps(source,indent=2)+'\n');rows=[]
    for name,command in checks:
      log=P/('quality-'+name+'.log');assert not log.exists()
      start=time.monotonic()
      with log.open('w') as stream:r=subprocess.run(command,cwd=ROOT,stdout=stream,stderr=subprocess.STDOUT)
      assert B.census()==source,'source changed during quality check'
      rows.append(dict(name=name,command=command,exit_code=r.returncode,seconds=time.monotonic()-start,log=log.name,log_sha256=B.sha(log),source_manifest=source_path.name,source_manifest_sha256=B.sha(source_path)))
      output.write_text(json.dumps(rows,indent=2)+'\n');print(name,r.returncode,flush=True)
      assert r.returncode==0
if __name__=='__main__':main()
