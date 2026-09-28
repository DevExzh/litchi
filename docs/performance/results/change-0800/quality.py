"""Root-only serial gates for the duplicate-error parity correction."""
import os, subprocess, time
import custody as c

p=c.P
out=p/'quality'
assert not out.exists()
out.mkdir()
source=c.source()
manifest=c.read(p/'correction/manifest.json')
base=c.read(p/'source.json')
changed={n:h for n,h in source['files'].items() if h!=base['files'][n]}
assert set(changed)==set(manifest['changed'])
for n,parts in manifest['changed'].items():
    assert c.sha(c.ROOT/n)==parts['after']['sha256']
c.write(out/'source.json',source)
inputs={str(f.relative_to(p)):c.sha(f) for folder in ['correction','edge-probe']
        for f in (p/folder).rglob('*') if f.is_file()}
inputs.update({n:c.sha(p/n) for n in ['quality.py','custody.py','origin.json','source.json','architecture-inputs.json','correction.patch','workspace-Cargo.lock']})
c.write(out/'inputs.json',inputs)
assert not os.environ.get('RUSTFLAGS')
env=os.environ|{'CARGO_TARGET_DIR':str(c.TARGET/'full'),'CARGO_BUILD_JOBS':'2','CARGO_INCREMENTAL':'0'}
packages=['litchi-opc','litchi-ole-common','litchi-sign','litchi-xldm','xml-minifier']
selectors=[arg for name in packages for arg in ['-p',name]]
commands=[
 ['rustfmt','--check','--edition','2024','--config','skip_children=true',*[str(c.ROOT/n) for n in manifest['changed']]],
 ['cargo','run','--offline','--locked','--manifest-path',str(p/'edge-probe/Cargo.toml')],
 ['cargo','test','--offline','--locked',*selectors,'--','--test-threads=2'],
 ['cargo','clippy','--offline','--locked',*selectors,'--all-targets','--','-D','warnings'],
]
rows=[]
for i,cmd in enumerate(commands):
    log=out/f'{i}.log'
    start=time.time()
    with log.open('w') as f:
        result=subprocess.run(cmd,cwd=c.ROOT,env=env,stdout=f,stderr=subprocess.STDOUT)
    rows.append({'command':cmd,'started':start,'ended':time.time(),'exit_code':result.returncode,'log':c.artifact(log)})
    c.write(out/'receipts.json',rows)
    assert result.returncode==0,log
    assert c.source()==source
    assert all(c.sha(p/n)==h for n,h in inputs.items())
    print('gate',i+1,'PASS',flush=True)
c.write(out/'complete.json',{'source':c.artifact(out/'source.json'),'inputs':c.artifact(out/'inputs.json'),'rows':rows,'scope':'Five affected full production crates, default features; controlled baseline/corrected edge oracle; no performance claim','environment':{n:env.get(n) for n in ['CARGO_TARGET_DIR','CARGO_BUILD_JOBS','CARGO_INCREMENTAL','RUSTFLAGS']}})
