"""Replay frozen source, capture/quality custody, reports and the optional seal."""
import argparse
import datetime
import hashlib
import json
import os
from pathlib import Path
import re
import subprocess
import tempfile

from run import HERE, FOLDER, REPO, TARGET, sha
from checks import COMMANDS
from analyze import read, receipt


def verify(sealed=False):
    plan=read(HERE/'plan.json')
    manifest=read(FOLDER/'source-manifest.json')
    for name,digest in manifest.items():
        assert sha(REPO/name)==digest,name
    review=(HERE/'final-source-review.md').read_text()
    assert manifest['tools/perf-baseline/src/lib.rs'] in review
    assert sha(FOLDER/'source.patch') in review
    for name,digest in read(HERE/'adr-manifest.json')['files'].items():
        assert sha(REPO/name)==digest,name
    with tempfile.TemporaryDirectory(prefix='litchi-0523-replay-') as directory:
        env=dict(os.environ,GIT_INDEX_FILE=str(Path(directory)/'index'))
        subprocess.run(['git','read-tree',plan['revision']],cwd=REPO,env=env,check=True)
        subprocess.run(['git','apply','--cached',str(FOLDER/'source.patch')],cwd=REPO,env=env,check=True)
        changed=subprocess.check_output(['git','diff','--cached','--name-only',plan['revision']],cwd=REPO,env=env,text=True).splitlines()
        assert changed==['tools/perf-baseline/src/lib.rs'],changed
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
    records=[]
    paths=sorted(FOLDER.glob('*.receipt.json'))
    assert len(paths)==30,len(paths)
    for path in paths:
        name=path.name.removesuffix('.receipt.json')
        kind=None if name.startswith(('build-','check-')) else ('alloc' if name.startswith('alloc-') else 'normal')
        r=receipt(name,kind,allow_failure=name.startswith('hardware-'))
        records.append(r)
    for name,command in COMMANDS:
        r=read(FOLDER/('check-'+name+'.receipt.json'))
        assert r['command']==['env','CARGO_TARGET_DIR='+str(TARGET),'CARGO_BUILD_JOBS=2',
            'CARGO_INCREMENTAL=0','RUSTDOCFLAGS=-D warnings']+command
    quality=read(HERE/'quality-summary.json')
    assert quality['status']=='pass' and len(quality['checks'])==10
    for row in quality['checks']:
        path=FOLDER/row['name'];assert sha(path)==row['receipt_sha256']
        log=path.with_name(path.name.replace('.receipt.json','.stdout')).read_text()
        assert row['executed_tests']==sum(map(int,re.findall(r'test result: ok\. (\d+) passed;',log)))
    assert quality['executed_tests']==sum(row['executed_tests'] for row in quality['checks'])
    preflight=read(HERE/'preflight-test.json')
    assert preflight['exit_code']==0 and preflight['source_unchanged']
    assert all(manifest[name]==digest for name,digest in preflight['source_hashes'].items())
    for name,digest in preflight['artifacts'].items():assert sha(HERE/name)==digest
    assert 'test result: ok. 1 passed;' in (HERE/'preflight-test.stdout').read_text()
    records.append(preflight)
    initial=HERE/'preflight-initial/preflight-test.json'
    if initial.exists():
        original=read(initial)
        assert original['exit_code']==0 and not original['source_unchanged']
        for name,digest in original['artifacts'].items():assert sha(initial.parent/name)==digest
        records.append(original)
    records.sort(key=lambda r:r['start_utc'])
    assert all(a['end_utc']<=b['start_utc'] for a,b in zip(records,records[1:]))
    native=sorted((read(p)['start_utc'],p.name.removesuffix('.receipt.json')) for p in FOLDER.glob('native-*.receipt.json'))
    assert [name for _,name in native]==['native-r1-xls','native-r1-cfb','native-r2-cfb','native-r2-xls']
    annotations={p:p.read_bytes() for p in FOLDER.glob('*.txt')}
    with tempfile.TemporaryDirectory(prefix='litchi-0523-analysis-') as directory:
        for script,name in [('analyze.py','analysis.json'),('analyze_profiles.py','profile-analysis.json'),('analyze_hardware.py','hardware-analysis.json')]:
            output=Path(directory)/name
            subprocess.run(['python3','-B',str(HERE/script),str(output)],cwd=REPO,check=True)
            assert output.read_bytes()==(HERE/name).read_bytes(),name
    assert annotations and all(path.read_bytes()==data for path,data in annotations.items())
    assert read(HERE/'analysis.json')['native_samples']==24000
    assert read(HERE/'analysis.json')['allocation_samples']==720
    negative=read(HERE/'verifier-tests.json')
    assert negative['status']=='pass' and len(negative['checks'])==4 and all(r['rejected'] for r in negative['checks'])
    flags=read(HERE/'variation-review.json');analysis=read(HERE/'analysis.json')
    assert flags['analysis_sha256']==sha(HERE/'analysis.json')
    assert len(flags['variations'])==len(analysis['same_build_variations_over_five_percent'])
    for raw,reviewed in zip(analysis['same_build_variations_over_five_percent'],flags['variations']):
        assert all(reviewed[k]==v for k,v in raw.items()) and reviewed['review']
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
    return dict(status='pass',native_samples=24000,allocation_samples=720,profile_children=8,
        production_unchanged=True,source_replay=True,serial_intervals=len(records),
        source_files=len(manifest),quality_gates=10,exact_report_replay=True,
        exact_annotation_replay=True,negative_vectors=4,owned_paths_absent=True,seal_entries=seal_count)


if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('--sealed',action='store_true');parser.add_argument('--output',type=Path)
    args=parser.parse_args();report=verify(args.sealed)
    if args.output:args.output.write_text(json.dumps(report,indent=2)+'\n')
    print(json.dumps(report))
