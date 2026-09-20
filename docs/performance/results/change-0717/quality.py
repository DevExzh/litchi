#!/usr/bin/env python3
"""Verify the changed standalone benchmark and its instrumentation guard."""
import importlib.util
import json
import os
from pathlib import Path
import subprocess
import time

P=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('custody0717quality',P/'custody.py')
C=importlib.util.module_from_spec(spec);spec.loader.exec_module(C)

def main():
    assert not (P/'quality.json').exists()
    source=C.census();assert source==json.loads((P/'source.json').read_text())
    support={n:C.sha(C.ROOT/n) for n in ['tools/test_perf_abba_summary.py','tools/perf-baseline/README.md']}
    (P/'support-files.json').write_text(json.dumps(support,indent=2)+'\n')
    base=['--release','--locked','--manifest-path','tools/perf-baseline/Cargo.toml','--target-dir',str(C.TARGET),'-j','2']
    tests=[
        ('fmt',['cargo','fmt','--manifest-path','tools/perf-baseline/Cargo.toml','--','--check'],{}),
        ('tests-native',['cargo','test',*base,'--lib'],{}),
        ('tests-procfs',['cargo','test',*base,'--lib','--features','ordinary-save-process-metrics'],{}),
        ('clippy',['cargo','clippy',*base,'--all-targets','--features','ordinary-save-process-metrics,allocator-metrics','--','-D','warnings'],{}),
        ('rustdoc',['cargo','doc',*base,'--no-deps','--lib','--features','ordinary-save-process-metrics'],{'RUSTDOCFLAGS':'-D warnings'}),
        ('latency-guard',['python3','-B','-m','unittest','discover','-s','tools','-p','test_perf_abba_summary.py'],{}),
    ]
    rows=[]
    for name,cmd,overrides in tests:
        env=dict(os.environ);env.update(overrides);log=P/f'quality-{name}.log';started=time.monotonic()
        with log.open('x') as out:r=subprocess.run(cmd,cwd=C.ROOT,env=env,stdout=out,stderr=subprocess.STDOUT)
        assert C.census()==source
        assert all(C.sha(C.ROOT/n)==h for n,h in support.items())
        rows.append(dict(name=name,command=cmd,environment_overrides=overrides,exit_code=r.returncode,
                         seconds=time.monotonic()-started,source_manifest_sha256=C.sha(P/'source.json'),
                         support_manifest_sha256=C.sha(P/'support-files.json'),log=log.name,log_sha256=C.sha(log)))
        (P/'quality.json').write_text(json.dumps(rows,indent=2)+'\n')
        print(name,'exit',r.returncode,flush=True);assert r.returncode==0

if __name__=='__main__':main()
