"""Strict evidence verification; no Rust build or capture mutation."""
from pathlib import Path
import datetime, gzip, hashlib, importlib.util, json, re, subprocess, tempfile
B=Path(__file__).resolve().parent
ROOT=B.parents[3]
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def read(n):return json.loads((B/n).read_text())
def verify():
    freeze=read('freeze.json')
    for n,h in freeze['sha256'].items():assert sha(B/n)==h,n
    for n,h in read('adr-manifest.json')['files'].items():assert sha(ROOT/n)==h,n
    original=(ROOT/'docs/performance/results/change-0544/candidate/sources/crates/litchi-xlsx/src/raw/worksheet/mod.rs').read_text()
    start=original.index('pub(crate) fn shared_event_bound_within_cap(')
    end=original.index('\n}',original.index('fn add_shared_event_bound(',start))+2
    original=original[start:end].replace('pub(crate) fn shared_event_bound_within_cap','pub fn baseline')
    actual=(B/'baseline.rs').read_text().replace('#[inline(never)]','')
    actual=actual[actual.index('pub fn baseline'):]
    assert re.sub(r'\s','',original)==re.sub(r'\s','',actual),'baseline adaptation'
    fixtures=read('fixtures.json')
    with tempfile.TemporaryDirectory(prefix='litchi-0546-verify-') as tmp:
        subprocess.run(['python3','-B',str(B/'generate.py'),tmp],cwd=ROOT,check=True)
        assert json.loads((Path(tmp)/'fixtures.json').read_text())==fixtures
        for f,item in fixtures.items():
            data=(Path(tmp)/item['file']).read_bytes()
            bound=1+bool(data and data[0] not in b'<&')+sum(c in b'<&' for c in data)+sum(c in b'>;' and n not in b'<&' for c,n in zip(data,data[1:]))
            for v in ['baseline','counted']:
                for repeat in [1,2]:
                    report=read(f'native-{v}-{repeat}-{f}.json')
                    assert report['bytes']==len(data) and report['full_bound']==bound
                    assert report['eligible']==(bound<=131072)
                    assert report['reader_events']<=bound
                    assert report['variant']==v and report['warmup']==10
                    assert len(report['duration_ns'])==report['samples']==100
                    assert all(isinstance(x,int) and x>0 for x in report['duration_ns'])
    names=['fmt','test','clippy','build']+[f'native-{v}-{f}' for v in freeze['order'] for f in freeze['fixtures']]+['assembly']
    assert {p.stem.removesuffix('.receipt') for p in B.glob('*.receipt.json')}==set(names)
    end=datetime.datetime.fromisoformat(freeze['utc']);binary_hashes=set()
    for name in names:
        r=read(name+'.receipt.json')
        assert r['freeze_sha256']==sha(B/'freeze.json') and r['exit_code']==0
        assert r['cwd']==str(ROOT)
        start=datetime.datetime.fromisoformat(r['started']);finish=datetime.datetime.fromisoformat(r['ended'])
        assert end<=start<=finish,name
        end=finish
        for stream in ['stdout','stderr']:
            path=B/(name+'.'+stream)
            data=path.read_bytes() if path.exists() else gzip.decompress(Path(str(path)+'.gz').read_bytes())
            assert hashlib.sha256(data).hexdigest()==r[stream+'_sha256']
        if 'binary_sha256' in r:binary_hashes.add(r['binary_sha256'])
        if name.startswith('native-'):
            v,rep,f=name.removeprefix('native-').split('-',2)
            assert r['argv']==['taskset','-c','2','/home/zhuhe/litchi-goal-0546-target/scanner/target/release/xlsx-scanner-diagnostic',v,'/home/zhuhe/litchi-goal-0546-target/fixtures/'+fixtures[f]['file'],str((B/(name+'.json')).relative_to(ROOT))]
    assert len(binary_hashes)==1
    assert '3 passed; 0 failed' in gzip.decompress((B/'test.stdout.gz').read_bytes()).decode()
    spec=importlib.util.spec_from_file_location('analysis545',B/'analyze.py');module=importlib.util.module_from_spec(spec);spec.loader.exec_module(module)
    summary=module.compute();assert summary==read('summary.json')
    assert len(summary['rows'])==20 and len(summary['reviews'])==10
    cleanup=read('cleanup.json');assert cleanup['binary_sha256'] in binary_hashes
    for n,h in cleanup['input_sha256'].items():assert h==freeze['sha256'][n]
    assert cleanup['removed'] and not Path(cleanup['target']).exists()
    assert datetime.datetime.fromisoformat(cleanup['utc'])>=end
    for path in B.glob('*.py'):compile(path.read_text(),str(path),'exec')
    seal=read('seal.json')
    actual_files={str(p.relative_to(ROOT)) for p in B.rglob('*') if p.is_file() and p.name!='seal.json'}
    assert actual_files=={n for n in seal['sha256'] if n.startswith(str(B.relative_to(ROOT))+'/')}
    for n,h in seal['sha256'].items():assert sha(ROOT/n)==h,n
    diff=subprocess.check_output(['git','diff',freeze['parent'],'--name-only'],cwd=ROOT,text=True).splitlines()
    decision=json.loads((B/'integration/decision.json').read_text())
    allowed={'crates/litchi-xlsx/examples/perf_cap_boundary.rs'}
    if decision['decision']=='retain':allowed.update(json.loads((B/'integration/plan.json').read_text())['candidate_files'])
    assert all(n.startswith('docs/performance/') or n in allowed for n in diff),diff
    subprocess.run(['git','diff','--check',freeze['parent'],'--','crates',':(glob)docs/performance/*.md','docs/performance/changes'],cwd=ROOT,check=True)
    subprocess.run(['python3','-B',str(B/'integration/verify.py'),'--strict'],cwd=ROOT,check=True)
    print('PASS: frozen diagnostic sources, exact baseline, regenerated fixtures, 45 serial receipts, 4000 samples, 10 reviews, cleanup, seal, integrated verifier')
if __name__=='__main__':verify()
