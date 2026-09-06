#!/usr/bin/env python3
"""Verify the sealed streaming evidence without its original binary or repo."""
import argparse
import gzip
import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
ROOT=Path(__file__).resolve().parent
def sha(raw): return hashlib.sha256(raw).hexdigest()
def path(name):
    p=(ROOT/name).resolve()
    assert p.is_relative_to(ROOT) and not Path(name).is_absolute(), name
    return p
def raw(name):
    p=path(name)
    return p.read_bytes() if p.is_file() else gzip.decompress(p.with_suffix(p.suffix+'.gz').read_bytes())
def load(name): return json.loads(raw(name))
def artifact(name, expected):
    data=raw(name); assert len(data)==expected['bytes'] and sha(data)==expected['sha256'],name

def verify():
    protocol=load('protocol.json'); build=load('build.json')
    assert set(build['binaries'])=={'normal','allocator'}
    assert sha(raw('protocol.json'))==build['protocol_sha256']
    assert sha(raw('verify-report.py'))==build['verifier_sha256']
    source=build['source_manifest']; assert sha(raw(source['path']))==source['sha256']
    assert len(load(source['path']))==source['files']
    for name, meta in build['build_receipts'].items():
        assert sha(raw(name))==meta['sha256']
        receipt=load(name); assert receipt['status']=='pass' and receipt['source_after']==source
    state=load('capture-state.json')
    assert state['status']=='pass' and state['tracked_tree_clean_before_and_after']
    assert state['status_before']==state['status_after']==['?? docs/GOAL.md','?? docs/performance/results/change-0432/']
    index=load('capture-index.json'); assert len(index)==len(protocol['order'])==12
    for name,lane in zip(index,protocol['order']):
        receipt=load(name); label='-'.join([lane['mode'],lane['shape'],lane['repeat'].lower()])
        assert receipt['name']==label and receipt['lane']==lane
        assert receipt['status']=='pass' and receipt['exit_code']==0
        assert receipt['revision']==build['revision'] and receipt['source_manifest']==source
        assert receipt['binary']==build['binaries'][lane['mode']]
        assert receipt['protocol_sha256']==build['protocol_sha256']
        assert receipt['driver_sha256']==sha(raw('capture.py'))
        assert receipt['verifier_sha256']==build['verifier_sha256']
        assert set(receipt['artifacts'])=={'captures/'+label+suffix for suffix in ['.json','-catalog.json','.log','-resource.log']}
        for target,meta in receipt['artifacts'].items(): artifact(target,meta)
        report=load('captures/'+label+'.json')
        assert report['environment']['git_revision']==build['revision']
        assert report['environment']['git_worktree_dirty'] is True
        assert report['binary_identity']['binary_sha256']==receipt['binary']['sha256']
        assert report['binary_identity']['binary_bytes']==receipt['binary']['bytes']
        catalog=load('captures/'+label+'-catalog.json')
        assert catalog['build']['git_revision']==build['revision']
        assert catalog['catalog_sha256']==report['corpus_catalog']['catalog_sha256']
        assert catalog['content_set_sha256']==report['corpus_catalog']['content_set_sha256']
        assert len(catalog['corpora'])==1 and catalog['corpora'][0]['legacy_v1']==report['results'][0]['corpus']
        argv=receipt['argv']; binary=receipt['binary']['path']
        offset=argv.index(binary)
        assert argv[:offset-1]==['taskset','-c',str(protocol['cpu']),'/usr/bin/time','-v','-o']
        tail=argv[offset+1:]; assert tail[:10]==['--case',protocol['selector'],'--semantic-shape',lane['shape'],'--workers','1','--samples',str(protocol['samples']),'--warmup',str(protocol['warmups'])]
        assert tail[10]=='--json' and Path(tail[11]).name==label+'.json'
        assert tail[12]=='--corpus-manifest' and Path(tail[13]).name==label+'-catalog.json' and len(tail)==14
        subprocess.run([sys.executable,'-B',str(ROOT/'verify-report.py'),'--report',str(ROOT/'captures'/(label+'.json')),'--mode',lane['mode'],'--shape',lane['shape']],check=True,stdout=subprocess.DEVNULL)
    checks=load('expected-checks.json')
    for tag,status in checks.items():
        r=load('checks/'+tag+'.json'); assert r['status']==status
        assert r['driver_sha256']==sha(raw('check.py'))
        for source_key in ['source_before','source_after']:
            s=r[source_key]; assert sha(raw(s['path']))==s['sha256']
        artifact(r['log']['path'],r['log'])
    for tag in load('planned-checks.json')['required_pass']:
        r=load('checks/'+tag+'.json'); assert r['status']=='pass' and r.get('exit_code',0)==0 and r.get('source_unchanged',True)
    for entry in (ROOT/'checks').glob('*.json'):
        r=json.loads(entry.read_text())
        if 'driver_hashes' in r:
            assert r['status']=='pass' and r['exit_code']==0 and r['drivers_unchanged'] and r['inventory_unchanged']
            for name,digest in r['driver_hashes'].items(): assert sha(raw(name))==digest
            artifact(r['log']['path'],r['log'])
    for name,r in load('compression.json').items():
        stored=path(name).read_bytes(); plain=gzip.decompress(stored)
        assert sha(stored)==r['stored_sha256'] and len(stored)==r['stored_bytes']
        assert sha(plain)==r['original_sha256'] and len(plain)==r['original_bytes']
    subprocess.run([sys.executable,'-B',str(ROOT/'derive.py'),'--check'],check=True,stdout=subprocess.DEVNULL)
    subprocess.run([sys.executable,'-B',str(ROOT/'compare-strict.py'),'--check'],check=True,stdout=subprocess.DEVNULL)
    inventory={line.split('  ',1)[1]:line.split('  ',1)[0] for line in raw('SHA256SUMS').decode().splitlines()}
    files={str(p.relative_to(ROOT)) for p in ROOT.rglob('*') if p.is_file() and p.name!='SHA256SUMS'}
    assert set(inventory)==files
    for name,digest in inventory.items(): assert sha(path(name).read_bytes())==digest,name
    return {'status':'pass','reports':12,'samples':360,'inventory_files':len(files),'command_receipts':len(checks)}

def portable():
    with tempfile.TemporaryDirectory(prefix='litchi-0432-portable-') as temp:
        exported=Path(temp)/'bundle'; shutil.copytree(ROOT,exported)
        subprocess.run([sys.executable,'-B',str(exported/'verify.py')],check=True,stdout=subprocess.DEVNULL)
        target=exported/'captures/allocator-tiny-r1.json'; original=target.read_bytes()
        mutations={
            'allocation_peak':lambda d:d['results'][0]['operation_metrics']['allocation']['region_peak_live_bytes']['values'].__setitem__(0,0),
            'output_identity':lambda d:d['results'][0].__setitem__('output_sha256','0'*64),
        }
        for name,mutate in mutations.items():
            doc=json.loads(original); mutate(doc); target.write_text(json.dumps(doc))
            result=subprocess.run([sys.executable,'-B',str(exported/'verify-report.py'),'--report',str(target),'--mode','allocator','--shape','tiny'],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
            assert result.returncode!=0,name
            target.write_bytes(original)
        verifier=exported/'verify-report.py'; verifier.write_bytes(verifier.read_bytes()+b'\n# mutated bound verifier\n')
        result=subprocess.run([sys.executable,'-B',str(exported/'verify.py')],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
        assert result.returncode!=0,'bound verifier mutation'
    return ['isolated export passed','allocation peak mutation rejected','output mutation rejected','bound verifier mutation rejected']
if __name__=='__main__':
    parser=argparse.ArgumentParser(); parser.add_argument('--portable-check',action='store_true'); args=parser.parse_args()
    result=verify()
    if args.portable_check: result['portable']=portable()
    print(json.dumps(result,indent=2))
