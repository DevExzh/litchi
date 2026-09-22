"""Freeze and capture the unchanged cross-copy lifecycle baseline."""
import copy, datetime, hashlib, json, pathlib, subprocess, sys, time
from build import P, ROOT, TARGET, census, sha, write
PHASES=['plan_ns','commit_ns','publication_ns','reopen_ns','lifecycle_ns']

def read(n): return json.loads((P/n).read_text())
def guard():
    assert census()==read('source.json')['files'],'source changed'
    for n,h in read('workspace-inputs.json')['files'].items(): assert sha(ROOT/n)==h,n
    for n,d in read('constraints.json').items(): assert sha(ROOT/n)==d,n

def schedule(qualification=False):
    plan=read('plan.json'); rows=[]
    for lane in ['native','allocation']:
        spec=plan[lane]
        for repeat in range(1 if qualification else spec['repeats']):
            for case in plan['cases'][::1 if repeat%2==0 else -1]:
                rows.append({'lane':lane,'repeat':repeat,'case':case,'samples':1 if qualification else spec['samples'],'warmups':0 if qualification else spec['warmups']})
    return rows

def command(row,out):
    b=next(b for b in read('build.json') if b['lane']==row['lane'])
    return ['taskset','-c',str(read('plan.json')['cpu']),b['binary'],'--warmup',str(row['warmups']),'--samples',str(row['samples']),'--case',row['case'],'--json',str(out)]

def projection(report,row):
    assert report['schema_version']==1
    assert len(report['results'])==1
    r=report['results'][0]; x=r['source']['pptx_cross_copy']; n=row['samples']
    assert r['case']==row['case']==r['corpus']['name']
    assert report['configuration']['samples_per_case']==n
    assert report['configuration']['warmup_iterations_per_case']==row['warmups']
    assert report['configuration']['cases']==[row['case']]
    b=next(b for b in read('build.json') if b['lane']==row['lane'])
    for k,v in [('path',b['binary']),('binary_sha256',b['binary_sha256']),('binary_bytes',b['binary_bytes'])]:
        assert report['binary_identity'][k]==v,k
    assert report['environment']['cpu_affinity']=='12'
    assert report['environment']['git_revision']==read('source.json')['head']
    if row['lane']=='allocation': assert report['tool']['allocator_counter_revision']=='serialized_region_peak_v3'
    assert report['tool']['profile']=='release'
    assert report['tool']['instrumentation']==('none' if row['lane']=='native' else 'system_allocator_operation_scoped')
    assert x['gates'] and all(v is True for v in x['gates'].values())
    assert r['output_sha256']==x['expected_output_sha256']
    assert len(x['output_sha256'])==n and all(v==r['output_sha256'] for v in x['output_sha256'])
    assert r['corpus']['archive_sha256']==x['destination_archive_sha256']
    elapsed=r['elapsed_ns']; assert len(elapsed['samples'])==n
    assert sorted(elapsed['sample_order'])==list(range(n))
    assert elapsed['samples']==sorted(elapsed['samples'])==x['lifecycle_ns']
    for phase in PHASES:
        assert len(x[phase])==n and all(type(v)is int and v>0 for v in x[phase]),phase
    for i,total in enumerate(x['lifecycle_ns']):
        assert sum(x[k][i] for k in ['plan_ns','commit_ns','publication_ns'])<=total
    alloc=r['operation_metrics']['allocation']
    assert alloc['status']==('unavailable' if row['lane']=='native' else 'measured')
    if row['lane']=='allocation':
        for k,v in alloc.items():
            if isinstance(v,dict):
                assert v['status']=='measured' and len(v['values'])==n and all(type(a)is int and a>=0 for a in v['values']),k
    stable={k:v for k,v in x.items() if k not in PHASES+['output_sha256']}
    return {'case':r['case'],'corpus':r['corpus'],'sink':r['sink'],'cross_copy':stable,'output_sha256':r['output_sha256']}

def capture(qualification=False):
    guard(); folder=P/('qualification' if qualification else 'captures'); folder.mkdir()
    rows=schedule(qualification); records=[]
    if not qualification: verify_freeze()
    for i,row in enumerate(rows):
        out=folder/f'{i:03d}.json'; cmd=command(row,out)
        b=next(b for b in read('build.json') if b['lane']==row['lane']); assert sha(pathlib.Path(b['binary']))==b['binary_sha256']
        start=time.time(); monotonic=time.monotonic_ns()
        with (folder/f'{i:03d}.stdout').open('w') as stdout, (folder/f'{i:03d}.stderr').open('w') as stderr:
            result=subprocess.run(cmd,cwd=ROOT,stdout=stdout,stderr=stderr)
        record=row|{'command':cmd,'started':start,'ended':time.time(),'monotonic_start':monotonic,'monotonic_end':time.monotonic_ns(),'exit':result.returncode,'files':{str(f.relative_to(P)):sha(f) for f in [out,folder/f'{i:03d}.stdout',folder/f'{i:03d}.stderr'] if f.exists()}}
        records.append(record); (folder/'manifest.json').write_text(json.dumps(records,indent=2)+'\n')
        assert result.returncode==0,record
        q=projection(json.loads(out.read_text()),row)
        if not qualification: assert q==read('oracle.json')[row['case']]
        print(f'PASS {folder.name} {i+1}/{len(rows)} {row["case"]} {row["lane"]}',flush=True)
    guard()
    if qualification:
        oracle={}
        for i,row in enumerate(rows):
            value=projection(json.loads((folder/f'{i:03d}.json').read_text()),row)
            if row['case'] in oracle: assert oracle[row['case']]==value
            oracle[row['case']]=value
        write('oracle.json',oracle)

def freeze():
    guard(); assert not (P/'freeze.json').exists()
    paths=['hypothesis.md','plan.json','source.json','workspace-inputs.json','constraints.json','environment.json','build.py','run.py','analyze.py','negative.py','negative.json','build.json','oracle.json','source-review.md']
    paths+= [str(p.relative_to(P)) for p in sorted((P/'qualification').glob('*'))]
    write('freeze.json',{'utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'files':{n:sha(P/n) for n in paths}})

def verify_freeze():
    for n,h in read('freeze.json')['files'].items(): assert sha(P/n)==h,n

if __name__=='__main__':
    {'qualify':lambda:capture(True),'freeze':freeze,'capture':capture}[sys.argv[1]]()
