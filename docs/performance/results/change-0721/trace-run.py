#!/usr/bin/env python3
"""Serial guarded source swaps, debug diagnostic builds, and exact public parity."""
import json,shutil,subprocess,sys,time
from custody import P,ROOT,TARGET,BIN,census,sha
lane=sys.argv[1];assert lane in ['baseline','candidate']
out=P/'trace'/lane;assert not out.exists();out.mkdir(parents=True)
candidate=json.loads((P/'source-candidate.json').read_text());baseline=json.loads((P/'source-baseline.json').read_text())
assert census()==candidate
changed=[n for n in sorted(candidate.keys()|baseline.keys()) if candidate.get(n)!=baseline.get(n)]
def install(label):
    expected=candidate if label=='candidate' else baseline
    for name in changed:
        target=ROOT/name
        if name not in expected:
            if target.exists():target.unlink()
        else:
            source=P/label/name;assert sha(source)==expected[name]
            shutil.copy2(source,target)
    assert census()==expected
try:
    if lane=='baseline':install('baseline')
    subprocess.run([sys.executable,str(P/'trace.py'),'apply',lane],cwd=ROOT,check=True)
    source=census();(out/'source.json').write_text(json.dumps(source,indent=2)+'\n')
    shutil.copy2(ROOT.parent/'litchi-scratch-0721-trace/manifest.json',out/'patch.json')
    command=['cargo','build','--locked','--manifest-path',str(P/'oracle/Cargo.toml'),'--target-dir',str(TARGET),'-j','2']
    start=time.monotonic()
    with (out/'build.log').open('w') as log:r=subprocess.run(command,cwd=ROOT,stdout=log,stderr=subprocess.STDOUT)
    assert census()==source
    record={'command':command,'exit_code':r.returncode,'seconds':time.monotonic()-start,'source_sha256':sha(out/'source.json'),'patch_sha256':sha(out/'patch.json'),'log_sha256':sha(out/'build.log'),'scripts':{n:sha(P/n) for n in ['trace-run.py','trace.py','trace.fragment','trace-analyze.py']}}
    (out/'build.json').write_text(json.dumps(record,indent=2)+'\n');assert r.returncode==0
    binary=BIN/(lane+'-trace');shutil.copy2(TARGET/'debug/docx-active-offset-oracle-0721',binary)
    record['binary']={'path':str(binary),'sha256':sha(binary),'bytes':binary.stat().st_size}
    record['probe']={n:sha(P/'oracle'/n) for n in ['Cargo.toml','Cargo.lock','src/main.rs']}
    (out/'build.json').write_text(json.dumps(record,indent=2)+'\n')
    for suffix in ['', '-repeat']:
        command=[sys.executable,str(P/'trace-analyze.py'),'capture','--binary',str(binary),'--report',str(out/('report'+suffix+'.json')),'--stdout',str(out/('stdout'+suffix)),'--stderr',str(out/('stderr'+suffix)),'--receipt',str(out/('capture'+suffix+'.json')),'--cwd',str(ROOT)]
        subprocess.run(command,cwd=ROOT,check=True)
        assert census()==source
        assert (out/('report'+suffix+'.json')).read_bytes()==(P/'oracle'/lane/'report.json').read_bytes()
    assert (out/'stderr').read_bytes()==(out/'stderr-repeat').read_bytes()
finally:
    scratch=ROOT.parent/'litchi-scratch-0721-trace'
    if scratch.exists():subprocess.run([sys.executable,str(P/'trace.py'),'restore'],cwd=ROOT,check=True)
    if lane=='baseline':install('candidate')
assert census()==candidate
print(lane,'trace build, repeated capture, public parity, and source restoration PASS',flush=True)
