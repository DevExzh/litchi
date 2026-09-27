"""Current-build custody and unchanged0739 semantic report contract."""
import hashlib, json, pathlib
from build import P, ROOT, TARGET, census, sha
PHASES=['plan_ns','commit_ns','publication_ns','reopen_ns','lifecycle_ns']
def read(n): return json.loads((P/n).read_text())
def guard():
    assert census()==read('source.json')['files']
    for n,h in read('workspace-inputs.json')['files'].items(): assert sha(ROOT/n)==h,n
    for n,h in read('constraints.json').items(): assert sha(ROOT/n)==h,n
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
