"""Validate retained per-sample public-export guardrail data and paired deltas."""
import csv
import hashlib
import json
import re
from pathlib import Path
from tools.summarize_crud_baseline import _rust_statistics

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def load(path):
    return json.loads(path.read_text())

def parse(rows, count=500):
    names={f'{pattern}-{length}' for pattern in ['dense_crlf','all_cr','sparse_crlf'] for length in [49,1024,65536]}
    assert len(rows)==9*count
    assert {r['case'] for r in rows}==names
    result={}
    for name in sorted(names):
        selected=[r for r in rows if r['case']==name]
        assert len(selected)==count
        assert [int(r['index']) for r in selected]==list(range(count))
        assert all(r['kind']=='sample' for r in selected)
        identities={(r['bytes'],r['objects'],r['archive_sha256'],r['output_sha256']) for r in selected}
        assert len(identities)==1
        identity=identities.pop()
        assert int(identity[0])>0 and int(identity[1])>0
        assert all(re.fullmatch('[0-9a-f]{64}',x) for x in identity[2:])
        samples=[int(r['elapsed_ns']) for r in selected]
        assert min(samples)>0
        order=sorted(range(count),key=lambda i:(samples[i],i))
        stats=_rust_statistics(samples,order)
        result[name]={'identity':identity,'statistics':stats}
    return result

def validate(here, count=500, prefix=""):
    root=here/'export-guardrail'
    reports={}
    for stage in ['before','after']:
        build=load(root/f'{stage}-build.json')
        assert build['exit_code']==0
        assert build['probe_source_sha256']==sha(root/'probe.rs')
        manifest=here/'before/source-manifest.json' if stage=='before' else here/'source-manifest.json'
        assert build['library_source_manifest_sha256']==sha(manifest)
        for repeat in [prefix+'r1',prefix+'r2']:
            receipt=load(root/f'{stage}-{repeat}-receipt.json')
            assert receipt['exit_code']==0 and receipt['source_unchanged']
            assert receipt['binary_sha256']==build['binary_sha256']
            assert receipt['build_receipt_sha256']==sha(root/f'{stage}-build.json')
            assert receipt['library_source_manifest_sha256']==sha(manifest)
            assert receipt['command'][:3]==['taskset','-c','2'] and receipt['command'][-1]==str(count)
            assert receipt['samples_per_case']==count and receipt['warmups']==10
            for name,digest in receipt['artifacts'].items():assert sha(root/name)==digest
            with (root/f'{stage}-{repeat}.csv').open() as stream:rows=list(csv.DictReader(stream))
            reports[stage,repeat]=parse(rows, count)
    for stage in ['before','after']:
        with (root/f'{stage}-preflight.csv').open() as stream:
            preflight=list(csv.DictReader(stream))
        assert len(preflight)==9
        identities={r['case']:(r['bytes'],r['objects'],r['archive_sha256'],r['output_sha256']) for r in preflight}
        assert len(identities)==9
        for repeat in [prefix+'r1',prefix+'r2']:
            for name,row in reports[stage,repeat].items():
                assert row['identity']==identities[name]
    order=[('before',prefix+'r1'),('after',prefix+'r1'),('after',prefix+'r2'),('before',prefix+'r2')]
    starts=[load(root/f'{stage}-{repeat}-receipt.json')['started_utc'] for stage,repeat in order]
    assert starts==sorted(starts) and len(set(starts))==4
    pairs=[]
    for repeat in [prefix+'r1',prefix+'r2']:
        for name in sorted(reports['before',repeat]):
            before,after=[reports[s,repeat][name] for s in ['before','after']]
            assert before['identity']==after['identity']
            b,a=before['statistics'],after['statistics']
            changes={m:(a[m]/b[m]-1)*100 for m in ['p50','mean','p95','p99']}
            changes['throughput']=(b['mean']/a['mean']-1)*100
            pairs.append({'repeat':repeat,'case':name,'before_p50_ns':b['p50'],'after_p50_ns':a['p50'],'changes_percent':changes,'adverse_over_5_percent':any(changes[m]>5 for m in ['p50','mean','p95','p99']) or changes['throughput'] < -5})
    drift=[]
    for stage in ['before','after']:
        for name in sorted(reports[stage,prefix+'r1']):
            a,b=[reports[stage,r][name]['statistics'] for r in [prefix+'r1',prefix+'r2']]
            changes={m:(b[m]/a[m]-1)*100 for m in ['p50','mean','p95','p99']}
            drift.append({'stage':stage,'case':name,'changes_percent':changes,'exceeds_drift_ceiling':any(abs(changes[m])>limit for m,limit in [('p50',5),('mean',5),('p95',10),('p99',15)])})
    rss=[]
    for repeat in [prefix+'r1',prefix+'r2']:
        b,a=[int(re.search(r'Maximum resident set size \(kbytes\): (\d+)',(root/f'{stage}-{repeat}.log').read_text())[1]) for stage in ['before','after']]
        rss.append({'repeat':repeat,'before_kib':b,'after_kib':a,'change_percent':(a/b-1)*100,'adverse_over_5_percent':a/b>1.05,'scope':'whole child across nine cases, fixture/open/preflight retained output included'})
    try:parse(rows[:-1], count)
    except AssertionError:pass
    else:raise AssertionError('short guardrail capture accepted')
    return {'samples':36*count,'pairs':pairs,'drift':drift,'whole_child_rss':rss,'negative_short_capture_rejected':True,'scope':'public export; counting discard sink, no timed hashing; separate from primary baseline'}
