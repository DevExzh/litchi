#!/usr/bin/env python3
"""Replay source-bound mapping/allocation attribution, without latency claims."""
import copy,hashlib,importlib.util,json,re,sys
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
def module(path,name):
    spec=importlib.util.spec_from_file_location(name,path);m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m);return m
def read(path):return json.loads(path.read_text())
def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest()
C=module(P/'custody.py','custody0718analysis')
CAP=module(P/'capture.py','capture0718analysis')
A=module(P.parent/'change-0717/analyze.py','prior0717analysis')
def validate_symbols(build):
    value=read(P/'symbolization.json');assert value['binary']==build['binary']
    assert value['script_sha256']==sha(P/'symbolize.py')
    assert value['addresses_sha256']==sha(P/'symbolization-addresses.txt')
    ref=value['source_reference'];assert ref['sha256']==sha(P/ref['file'])
    for row in value['commands']:assert row['exit_code']==0 and row['sha256']==sha(P/row['stdout'])
    for name,h in value['traces'].items():assert sha(P/name)==h
    assert set(value['traces'])=={f.name for f in P.glob('mapping-r*.strace')}
    pattern=re.compile(re.escape(value['binary']['path'])+r'\([^\n]*\) \[(0x[0-9a-fA-F]+)\]')
    offsets={hex(int(x,16)) for name in value['traces'] for x in pattern.findall((P/name).read_text())}
    assert offsets==set(value['frames'])
    segments=[]
    for line in (P/'symbolization-program-headers.txt').read_text().splitlines():
        m=re.fullmatch(r'\s*LOAD\s+(0x[0-9a-f]+)\s+(0x[0-9a-f]+)\s+(0x[0-9a-f]+)\s+(0x[0-9a-f]+)\s+(0x[0-9a-f]+)\s+([RWE ]+)\s+(0x[0-9a-f]+)\s*',line)
        if m and 'E' in m[6]:segments.append(dict(file_offset=int(m[1],16),virtual_address=int(m[2],16),file_size=int(m[4],16),memory_size=int(m[5],16),flags=sum(bit for c,bit in [('R',4),('W',2),('E',1)] if c in m[6]),alignment=int(m[7],16)))
    assert segments==value['segments']
    symbols=[]
    for line in (P/'symbolization-symbols.txt').read_text().splitlines():
        m=re.fullmatch(r'([0-9a-f]+) ([0-9a-f]+) ([TtWw]) (.+)',line)
        if m and int(m[2],16)>0:symbols.append(dict(address=int(m[1],16),bytes=int(m[2],16),name=m[4]))
    resolved=(P/'symbolization-resolved.txt').read_text().splitlines();assert len(resolved)==2*len(offsets)
    addresses=[]
    for i,(key,frame) in enumerate(value['frames'].items()):
        offset=int(key,16);assert frame['file_offset']==offset
        segments=[s for s in value['segments'] if s['file_offset']<=offset<s['file_offset']+s['file_size']];assert len(segments)==1
        segment=segments[0];addr=offset-segment['file_offset']+segment['virtual_address'];assert addr==frame['virtual_address']
        matches=[s for s in symbols if s['address']<=addr<s['address']+s['bytes']]
        assert matches==frame['covering_symbols'] and any(s['name']==frame['function'] for s in matches)
        assert resolved[2*i:2*i+2]==[frame['function'],frame['location']]
        addresses.append(hex(addr))
    assert (P/'symbolization-addresses.txt').read_text()==''.join(x+'\n' for x in addresses)
    return value

def analyze():
    T=module(P/'attribution.py','attribution0718')
    plan,build,source=[read(P/n) for n in ['plan.json','build.json','source.json']]
    assert source==C.census()==read(P.parent/'change-0717/source.json'),'source changed'
    for filename in ['constraints.json','helper-freeze.json']:
        for n,h in read(P/filename).items():assert sha(ROOT/n)==h,(filename,n)
    for n,h in read(P/'capture-freeze.json').items():assert sha(P/n)==h,('capture freeze',n)
    assert build['exit_code']==0 and build['source_sha256']==sha(P/'source.json') and build['log_sha256']==sha(P/'build.log')
    expected_build=['cargo','build','--release','--locked','--manifest-path','tools/perf-baseline/Cargo.toml','--bin','litchi-perf-baseline','--target-dir',str(C.TARGET),'-j','2','--features','ordinary-save-process-metrics']
    assert build['command']==expected_build
    binary=build['binary'];path=Path(binary['path'])
    assert path==C.BIN/'procfs'
    if path.exists():assert sha(path)==binary['sha256'] and path.stat().st_size==binary['bytes']
    else:assert read(P/'cleanup.json')['binaries']==[binary],'missing binary cleanup witness'
    symbols=validate_symbols(build)
    expected=CAP.jobs(plan);assert len(expected)==8
    assert {f.name for f in P.glob('*.receipt.json')}=={j['name']+'.receipt.json' for j in expected}
    rows=[];last_started=None
    for job in expected:
        name=job['name'];receipt=read(P/(name+'.receipt.json'))
        assert receipt['job']==job and receipt['exit_code']==0 and receipt['command']==CAP.command(plan,job,build),(name,'job/command')
        assert receipt['source_before']==receipt['source_after']==CAP.digest(source)
        for key,file in [('source_manifest_sha256','source.json'),('build_sha256','build.json'),('plan_sha256','plan.json'),('script_sha256','capture.py'),('tools_sha256','tool-identities.json')]:assert receipt[key]==sha(P/file),(name,key)
        assert receipt['binary']==binary and receipt['environment']==plan['expected_environment']
        assert receipt['fixture_before']==receipt['fixture_after']==CAP.fixture(job['corpus'])
        assert isinstance(receipt['seconds'],(float,int)) and receipt['seconds']>0
        if last_started is not None:assert receipt['started_utc']>last_started
        last_started=receipt['started_utc']
        suffixes=['json','stdout','stderr', 'strace' if job['lane']=='mapping' else 'dhat.json']
        files={name+'.'+suffix for suffix in suffixes};assert set(receipt['artifacts'])==files
        assert {f.name for f in P.glob(name+'.*')}==files|{name+'.receipt.json'}
        for n,h in receipt['artifacts'].items():assert sha(P/n)==h,(name,n)
        assert (P/(name+'.stdout')).read_bytes()==b''
        report=read(P/(name+'.json'))
        check_plan=dict(plan,samples=job['samples'],warmup=job['warmup'])
        result,elapsed,execution,parity,process=A.validate_report(check_plan,job,build,report,sha(P/'source.json'))
        # Validate every process vector and raw delta, but do not duplicate them or
        # reinterpret DHAT/strace procfs observations as native memory evidence.
        if job['lane']=='mapping':attribution=T.parse_mapping(P/(name+'.strace'),plan,report,symbols)
        else:attribution=T.parse_allocation(P/(name+'.dhat.json'),plan)
        rows.append(dict(job=job,receipt_sha256=sha(P/(name+'.receipt.json')),artifacts=receipt['artifacts'],normalized_parity=parity,attribution=attribution))
    return dict(schema_version=1,revision=plan['revision'],scope='Unchanged-source stack attribution; no native timing, hardware fault cause, allocator policy, RSS or speedup claim.',source_sha256=sha(P/'source.json'),build_sha256=sha(P/'build.json'),binary=binary,rows=rows,all_children_retained=True,all_output_parity_verified=True,claims=plan['claims'],bindings={n:sha(P/n) for n in ['plan.json','capture.py','attribution.py','analyze.py','capture-freeze.json','helper-freeze.json','tool-identities.json','symbolization.json','symbolize.py']})
def main():
    value=analyze();path=P/'analysis.json'
    if sys.argv[1:]==['--check']:assert value==read(path),'derived analysis differs';print('PASS exact0718 replay,8 children')
    else:
        assert not path.exists(),'refusing derived overwrite';path.write_text(json.dumps(value,indent=2)+'\n');print('PASS0718 attribution,8 children')
if __name__=='__main__':main()
