#!/usr/bin/env python3
"""Build unchanged-source procfs diagnostic executable for stack attribution."""
import importlib.util,json,shutil,subprocess,time
from pathlib import Path
P=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('custody0718',P/'custody.py');C=importlib.util.module_from_spec(spec);spec.loader.exec_module(C)
def main():
    assert not (P/'build.json').exists()
    source=C.census();assert source==json.loads((P/'source.json').read_text())==json.loads((P.parent/'change-0717/source.json').read_text())
    command=['cargo','build','--release','--locked','--manifest-path','tools/perf-baseline/Cargo.toml','--bin','litchi-perf-baseline','--target-dir',str(C.TARGET),'-j','2','--features','ordinary-save-process-metrics']
    start=time.monotonic()
    with (P/'build.log').open('x') as out:r=subprocess.run(command,cwd=C.ROOT,stdout=out,stderr=subprocess.STDOUT)
    assert C.census()==source
    row=dict(command=command,exit_code=r.returncode,seconds=time.monotonic()-start,source_sha256=C.sha(P/'source.json'),log_sha256=C.sha(P/'build.log'))
    if r.returncode==0:
        C.BIN.mkdir();binary=C.BIN/'procfs';shutil.copy2(C.TARGET/'release/litchi-perf-baseline',binary);row['binary']=dict(path=str(binary),sha256=C.sha(binary),bytes=binary.stat().st_size)
    (P/'build.json').write_text(json.dumps(row,indent=2)+'\n');print('diagnostic build',r.returncode,flush=True);assert r.returncode==0
if __name__=='__main__':main()
