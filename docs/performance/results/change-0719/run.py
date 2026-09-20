#!/usr/bin/env python3
"""Serial source-bound build, correctness captures and quality gates."""
import importlib.util,json,os,subprocess,sys,time
from pathlib import Path
P=Path(__file__).resolve().parent
s=importlib.util.spec_from_file_location('custody0719',P/'custody.py');C=importlib.util.module_from_spec(s);s.loader.exec_module(C)
def run(name,command,env=None):
    source=C.census();assert source==json.loads((P/'source.json').read_text())
    log=P/(name+'.log'); receipt=P/(name+'.receipt.json');assert not log.exists() and not receipt.exists()
    started=time.monotonic()
    with log.open('x') as f:r=subprocess.run(command,cwd=C.ROOT,env=dict(os.environ,**(env or {})),stdout=f,stderr=subprocess.STDOUT)
    assert C.census()==source
    row=dict(command=command,environment_overrides=env or {},exit_code=r.returncode,seconds=time.monotonic()-started,source_sha256=C.sha(P/'source.json'),log_sha256=C.sha(log),runner_sha256=C.sha(Path(__file__)))
    receipt.write_text(json.dumps(row,indent=2)+'\n');print(name,r.returncode,flush=True);assert r.returncode==0
base=['--release','--locked','--manifest-path','tools/perf-baseline/Cargo.toml','--target-dir',str(C.TARGET),'-j','2']
if sys.argv[1:]==['build']:
    run('build',['cargo','build',*base,'--bin','litchi-perf-baseline'])
    binary=C.TARGET/'release/litchi-perf-baseline'
    (P/'binary.json').write_text(json.dumps(dict(path=str(binary),sha256=C.sha(binary),bytes=binary.stat().st_size),indent=2)+'\n')
elif sys.argv[1:]==['quality']:
    run('fmt',['cargo','fmt','--manifest-path','tools/perf-baseline/Cargo.toml','--','--check'])
    run('tests',['cargo','test',*base,'--lib'])
    run('clippy',['cargo','clippy',*base,'--all-targets','--features','ordinary-save-process-metrics,allocator-metrics','--','-D','warnings'])
    run('rustdoc',['cargo','doc',*base,'--no-deps','--lib'],{'RUSTDOCFLAGS':'-D warnings'})
elif sys.argv[1:]==['capture']:
    b=json.loads((P/'binary.json').read_text());assert C.sha(Path(b['path']))==b['sha256']
    for repeat,shapes in [(1,['medium','dense']),(2,['dense','medium'])]:
        for shape in shapes:
            name=f'correctness-r{repeat}-{shape}'
            command=[b['path'],'--warmup','1','--samples','3','--case',f'xlsx_producer_{shape}_source_one_edit_save','--producer-evidence',str(P/(name+'.corpus.json')),'--json',str(P/(name+'.json'))]
            run(name,command)
    assert C.sha(Path(b['path']))==b['sha256']
elif sys.argv[1:]==['evidence']:
    for r in json.loads((P.parent/'change-0717/evidence/results.json').read_text()):run('gate-'+r['name'],r['command'])
else:raise SystemExit('usage: run.py build|quality|capture|evidence')
