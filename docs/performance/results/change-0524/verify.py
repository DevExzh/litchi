"""Verify matched CFB evidence custody, source replay, admission and seal."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import subprocess
import tempfile

import analyze as numeric
from checks import COMMANDS
from run import HERE, REPO, TARGET, sha


def read(path):
    return json.loads(path.read_text())


def replay_source(stage, plan):
    FOLDER=HERE/stage
    manifest=read(FOLDER/'source-manifest.json')
    with tempfile.TemporaryDirectory(prefix='litchi-0523-replay-') as directory:
        env=dict(os.environ,GIT_INDEX_FILE=str(Path(directory)/'index'))
        subprocess.run(['git','read-tree',plan['revision']],cwd=REPO,env=env,check=True)
        if (FOLDER/'source.patch').stat().st_size:
            subprocess.run(['git','apply','--cached',str(FOLDER/'source.patch')],cwd=REPO,env=env,check=True)
        else:
            assert stage=='baseline'
        changed=subprocess.check_output(['git','diff','--cached','--name-only',plan['revision']],cwd=REPO,env=env,text=True).splitlines()
        assert changed==([] if stage=='baseline' else (['crates/litchi-cfb/src/file.rs','crates/litchi-cfb/src/writer/sequential.rs'] if stage=='final' else ['crates/litchi-cfb/src/file.rs'])),changed
        entries=subprocess.check_output(['git','ls-files','-s','-z'],cwd=REPO,env=env)
        objects={}
        for entry in entries.split(b'\0'):
            if entry:
                metadata,name=entry.split(b'\t',1)
                objects[name.decode()]=metadata.split()[1].decode()
        unique={objects[name] for name in manifest}
        batch=subprocess.check_output(['git','cat-file','--batch'],cwd=REPO,input=('\n'.join(sorted(unique))+'\n').encode())
        hashes={};position=0
        while position<len(batch):
            end=batch.index(b'\n',position);oid,kind,size=batch[position:end].split();assert kind==b'blob'
            position=end+1;data=batch[position:position+int(size)]
            hashes[oid.decode()]=hashlib.sha256(data).hexdigest();position+=int(size)+1
        assert all(hashes[objects[name]]==digest for name,digest in manifest.items())

def verify(sealed=False):
    plan=read(HERE/'plan.json')
    decision=read(HERE/'decision.json')
    assert decision['disposition'] in ('accepted','rejected')
    current='candidate' if decision['disposition']=='accepted' else 'final'
    for name,digest in read(HERE/current/'source-manifest.json').items():
        assert sha(REPO/name)==digest,name
    if current=='final':
        final_review=(HERE/'final-review.md').read_text()
        assert sha(HERE/'final/source-manifest.json') in final_review and sha(HERE/'final/source.patch') in final_review
        for name in ['crates/litchi-cfb/src/file.rs','crates/litchi-cfb/src/writer/sequential.rs','tools/perf-baseline/src/lib.rs']:
            assert read(HERE/'final/source-manifest.json')[name] in final_review
    for name,digest in read(HERE/'adr-manifest.json')['files'].items():
        assert sha(REPO/name)==digest,name
    stages=['baseline','candidate']+(['final'] if current=='final' else [])
    source_review=(HERE/'source-review.md').read_text()
    for stage in ('baseline','candidate'):
        assert read(HERE/stage/'source-manifest.json')['crates/litchi-cfb/src/file.rs'] in source_review
    for name in ['candidate.patch','tests.patch','collector-tests.patch','preparation-adjustment.json','preflight-environment-adjustment.json']:
        assert sha(HERE/name) in source_review
    records=[]
    for stage in stages:
        replay_source(stage,plan)
        numeric.configure(stage)
        for path in sorted((HERE/stage).glob('*.receipt.json')):
            name=path.name.removesuffix('.receipt.json')
            kind=None if name.startswith(('build-','check-')) else ('alloc' if name.startswith('alloc-') else 'normal')
            records.append((stage,name,numeric.receipt(name,kind,allow_failure=name.startswith('hardware-') or (stage=='candidate' and name=='check-cfb-tests'))))
    assert len(records)==(57 if current=='final' else 54),len(records)
    if current=='final':
        failed=read(HERE/'candidate/check-cfb-tests.receipt.json')
        assert failed['exit_code']==101
        assert '277 passed; 1 failed;' in (HERE/'candidate/check-cfb-tests.stdout').read_text()
        for name in ['crates/litchi-cfb/src/file.rs','crates/litchi-cfb/src/writer/sequential.rs']:
            original=subprocess.check_output(['git','show',plan['revision']+':'+name],cwd=REPO)
            assert original.split(b'#[cfg(test)]',1)[0]==(REPO/name).read_bytes().split(b'#[cfg(test)]',1)[0]
        probe=read(HERE/'temp-identity-probe.json')
        assert probe['reused']['unlink']>0 and probe['reused']['rename']==0 and probe['owned_probe_directory_absent']
    for name,command in COMMANDS:
        r=read(HERE/current/('check-'+name+'.receipt.json'))
        assert r['command']==['env','TMPDIR='+str(TARGET/'test-tmp'),'CARGO_TARGET_DIR='+str(TARGET),'CARGO_BUILD_JOBS=2',
            'CARGO_INCREMENTAL=0','RUSTDOCFLAGS=-D warnings']+command
    quality=read(HERE/'quality-summary.json')
    assert quality['status']=='pass' and len(quality['checks'])==14 and quality['stage']==current
    for row in quality['checks']:
        path=HERE/quality['stage']/row['name'];assert sha(path)==row['receipt_sha256']
        log=path.with_name(path.name.replace('.receipt.json','.stdout')).read_text()
        assert row['executed_tests']==sum(map(int,re.findall(r'test result: ok\. (\d+) passed;',log)))
    assert quality['executed_tests']==sum(row['executed_tests'] for row in quality['checks'])
    preflight=read(HERE/'preflight.receipt.json')
    assert preflight['exit_code']==0 and preflight['source_unchanged']
    assert preflight['source_manifest_sha256']==sha(HERE/'preflight-source-manifest.json')
    assert read(HERE/'preflight-source-manifest.json')==read(HERE/'candidate/source-manifest.json')
    for name,digest in preflight['artifacts'].items():assert sha(HERE/name)==digest
    assert re.search(r'test result: ok\. [1-9][0-9]* passed;', (HERE/'preflight.stdout').read_text())
    records.append(('preflight','preflight',preflight))
    initial=read(HERE/'preflight-initial/preflight.receipt.json')
    assert initial['exit_code']==101 and initial['source_unchanged']
    assert initial['source_manifest_sha256']==sha(HERE/'preflight-initial/preflight-source-manifest.json')
    assert read(HERE/'preflight-initial/preflight-source-manifest.json')==read(HERE/'preflight-source-manifest.json')
    for name,digest in initial['artifacts'].items():assert sha(HERE/'preflight-initial'/name)==digest
    records.append(('preflight','preflight-initial',initial))
    records.sort(key=lambda r:r[2]['start_utc'])
    assert all(a[2]['end_utc']<=b[2]['start_utc'] for a,b in zip(records,records[1:]))
    native=[(stage,name) for stage,name,r in records if name.startswith('native-')]
    assert native==[(s,'native-r'+str(r)+'-'+g) for s,r,groups in
        [('baseline',1,['xls','cfb']),('candidate',1,['xls','cfb']),
         ('candidate',2,['cfb','xls']),('baseline',2,['cfb','xls'])] for g in groups]
    for stage,name,r in records:
        if stage=='preflight':continue
        expected='candidate' if stage=='baseline' and name.startswith('native-r2-') else stage
        assert r['execution_stage']==expected,(stage,name)
    annotations={p:p.read_bytes() for stage in ('baseline','candidate') for p in (HERE/stage).glob('*.txt')}
    with tempfile.TemporaryDirectory(prefix='litchi-0524-analysis-') as directory:
        for stage in ('baseline','candidate'):
            for script,name in [('analyze.py','analysis.json'),('analyze_profiles.py','profile-analysis.json'),('analyze_hardware.py','hardware-analysis.json')]:
                output=Path(directory)/(stage+'-'+name)
                subprocess.run(['python3','-B',str(HERE/script),'--stage',stage,str(output)],cwd=REPO,check=True)
                assert output.read_bytes()==(HERE/stage/name).read_bytes(),(stage,name)
        for script,name in [('analyze.py','comparison.json'),('analyze_profiles.py','profile-comparison.json')]:
            output=Path(directory)/name
            subprocess.run(['python3','-B',str(HERE/script),'--compare',str(output)],cwd=REPO,check=True)
            assert output.read_bytes()==(HERE/name).read_bytes(),name
    assert len(annotations)==160 and all(path.read_bytes()==data for path,data in annotations.items())
    for stage in ('baseline','candidate'):
        a=read(HERE/stage/'analysis.json')
        assert a['native_samples']==24000 and a['allocation_samples']==720
    negative=read(HERE/'verifier-tests.json')
    assert negative['status']=='pass' and len(negative['checks'])==4 and all(r['rejected'] for r in negative['checks'])
    from decide import evaluate
    assert decision==evaluate()
    assert decision['comparison_sha256']==sha(HERE/'comparison.json')
    for stage in ('baseline','candidate'):
        assert decision['profiles_sha256'][stage]==sha(HERE/stage/'profile-analysis.json')
    before=read(HERE/'baseline/analysis.json');after=read(HERE/'candidate/analysis.json')
    native_rows=lambda a:{(r['case'],r['shape'],r['repeat']):r for r in a['rows'] if r['lane']=='native'}
    a,b=native_rows(before),native_rows(after)
    primary=[]
    for key,x in a.items():
        if key[0] in plan['review']['primary_cases']:
            primary.append(100*(b[key]['elapsed_ns']['p50']/x['elapsed_ns']['p50']-1))
    native_gate=len(primary)==8 and all(x<=-3 for x in primary)
    assert decision['native_primary_gate']==native_gate
    if decision['disposition']=='accepted':
        assert native_gate and decision['profile_gate'] and decision['memory_gate'] and decision['adverse_review_complete']
    else:
        assert not native_gate or not decision['profile_gate'] or not decision['memory_gate'] or decision['other_rejection_reason']
    cleanup=read(HERE/'cleanup.json')
    assert cleanup['owned_paths_absent'] and cleanup['python_cache_absent'] and not cleanup['accessible_process_references']
    assert all(not Path(name).exists() for name in plan['owned_paths'])
    assert not list(HERE.rglob('__pycache__'))
    seal_count=None
    if sealed:
        expected={}
        for line in (HERE/'SHA256SUMS').read_text().splitlines():
            digest,name=line.split('  ',1)
            assert name not in expected and not Path(name).is_absolute() and '..' not in Path(name).parts and name!='SHA256SUMS'
            expected[name]=digest
        actual={str(p.relative_to(HERE)):sha(p) for p in HERE.rglob('*') if p.is_file() and p.name!='SHA256SUMS'}
        assert expected==actual
        assert not any(p.is_symlink() for p in HERE.rglob('*'))
        seal_count=len(expected)
    return dict(status='pass',disposition=decision['disposition'],native_samples=48000,allocation_samples=1440,
        profile_children=16,source_replay=True,serial_intervals=len(records),quality_gates=14,
        exact_report_replay=True,exact_annotation_replay=True,negative_vectors=4,owned_paths_absent=True,seal_entries=seal_count)


if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('--sealed',action='store_true');parser.add_argument('--output',type=Path)
    args=parser.parse_args();report=verify(args.sealed)
    if args.output:args.output.write_text(json.dumps(report,indent=2)+'\n')
    print(json.dumps(report))
