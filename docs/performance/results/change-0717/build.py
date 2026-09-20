#!/usr/bin/env python3
"""Freeze native and explicitly instrumented harness binaries from one source."""
import importlib.util
import json
from pathlib import Path
import shutil
import subprocess
import time

P=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('custody0717build',P/'custody.py')
C=importlib.util.module_from_spec(spec);spec.loader.exec_module(C)

def main():
    assert not (P/'builds.json').exists()
    source=C.census();baseline=json.loads((P/'source-baseline.json').read_text())
    changed={n for n in set(source)|set(baseline) if source.get(n)!=baseline.get(n)}
    assert changed=={'tools/perf-baseline/Cargo.toml','tools/perf-baseline/src/lib.rs','tools/perf-baseline/src/ordinary_save.rs'},changed
    (P/'source.json').write_text(json.dumps(source,indent=2)+'\n')
    C.BIN.mkdir();rows={}
    for lane,args in [('native',[]),('procfs',['--features','ordinary-save-process-metrics'])]:
        command=['cargo','build','--release','--locked','--manifest-path','tools/perf-baseline/Cargo.toml',
                 '--bin','litchi-perf-baseline','--target-dir',str(C.TARGET),'-j','2',*args]
        log=P/f'build-{lane}.log';started=time.monotonic()
        with log.open('x') as out:result=subprocess.run(command,cwd=C.ROOT,stdout=out,stderr=subprocess.STDOUT)
        assert C.census()==source
        row=dict(command=command,exit_code=result.returncode,seconds=time.monotonic()-started,
                 source_sha256=C.sha(P/'source.json'),log=log.name,log_sha256=C.sha(log))
        if result.returncode==0:
            binary=C.BIN/lane;shutil.copy2(C.TARGET/'release/litchi-perf-baseline',binary)
            row['binary']=dict(path=str(binary),sha256=C.sha(binary),bytes=binary.stat().st_size)
        rows[lane]=row;(P/'builds.json').write_text(json.dumps(rows,indent=2)+'\n')
        print(lane,'build exit',result.returncode,flush=True);assert result.returncode==0

if __name__=='__main__':main()
