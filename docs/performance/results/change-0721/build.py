#!/usr/bin/env python3
import importlib.util,json,os,shutil,subprocess,sys,time
from pathlib import Path
P=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('custody0721',P/'custody.py');C=importlib.util.module_from_spec(spec);spec.loader.exec_module(C)
def main():
    stage=sys.argv[1];assert stage in ['baseline','candidate'];dest=P/('build-'+stage+'.json');assert not dest.exists()
    source=C.census();manifest=P/('source-'+stage+'.json');manifest.write_text(json.dumps(source,indent=2)+'\n');C.BIN.mkdir(exist_ok=True)
    rows=[]
    for lane,name,args in [('native','litchi-perf-baseline',[]),('alloc','litchi-perf-baseline-alloc',['--features','allocator-metrics'])]:
        cmd=['cargo','build','--release','--locked','--manifest-path','tools/perf-baseline/Cargo.toml','--bin',name,'--target-dir',str(C.TARGET),'-j','2',*args];log=P/f'build-{stage}-{lane}.log';start=time.monotonic()
        with log.open('w') as f:r=subprocess.run(cmd,cwd=C.ROOT,stdout=f,stderr=subprocess.STDOUT)
        assert C.census()==source
        row=dict(lane=lane,command=cmd,exit_code=r.returncode,seconds=time.monotonic()-start,log_sha256=C.sha(log),source_manifest_sha256=C.sha(manifest))
        if r.returncode==0:
            binary=C.BIN/f'{stage}-{lane}';shutil.copy2(C.TARGET/'release'/name,binary);row.update(binary=str(binary),binary_sha256=C.sha(binary),binary_bytes=binary.stat().st_size)
        rows.append(row);dest.write_text(json.dumps(rows,indent=2)+'\n');print(stage,lane,r.returncode,flush=True);assert r.returncode==0
if __name__=='__main__':main()
