#!/usr/bin/env python3
"""Verify final source/raw bindings and compare every recorded 0683 workload."""
import hashlib
import json
from pathlib import Path
import statistics

PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]

def digest(p): return hashlib.sha256(p.read_bytes()).hexdigest()
def row(line): return dict(field.split('=',1) for field in line.strip().split('\t'))
def stats(path):
    rows=[row(line) for line in path.read_text().splitlines()]
    assert len(rows)==20, path
    assert [int(r['sample']) for r in rows]==list(range(20)),path
    semantics={(r['result'],r['callbacks']) for r in rows}
    assert len(semantics)==1,path
    values=sorted(int(r['nanos']) for r in rows)
    return dict(p50=statistics.median(values),mean=statistics.mean(values),p95=values[18],p99=values[19]), next(iter(semantics))
def delta(after,before): return (after/before-1)*100 if before else None

manifests={}
for phase in ['baseline','candidate']:
    folder=PACKET/'measurements'/phase
    manifest=json.loads((folder/'manifest.json').read_text());manifests[phase]=manifest
    for name,expected in manifest['raw_sha256'].items(): assert digest(folder/name)==expected,(phase,name)
    for name,expected in manifest['probe_sha256'].items(): assert digest(PACKET/name)==expected,(phase,name)
assert manifests['baseline']['probe_sha256']==manifests['candidate']['probe_sha256']
assert manifests['baseline']['corpus']==manifests['candidate']['corpus']
for name,expected in manifests['candidate']['source_sha256'].items(): assert digest(ROOT/name)==expected,name
constraints=json.loads((PACKET/'baseline.json').read_text())
for name,expected in constraints['constraints_sha256'].items(): assert digest(ROOT/name)==expected,name
quality=json.loads((PACKET/'final-verified/results.json').read_text())
assert {r['name'] for r in quality}=={'fmt','check','clippy','tests','facade','rustdoc'}
for command in quality:
    assert command['exit_code']==0
    for name,expected in command['source_sha256'].items(): assert digest(ROOT/name)==expected,name

# Archive paths changed, but raw and probe hashes remain bound to their capture.
archive=PACKET/'initial-candidate'
for phase in ['baseline','candidate']:
    folder=archive/'measurements'/phase
    manifest=json.loads((folder/'manifest.json').read_text())
    for name,expected in manifest['raw_sha256'].items(): assert digest(folder/name)==expected,('initial',phase,name)
    for name,expected in manifest['probe_sha256'].items(): assert digest(archive/name)==expected,('initial',phase,name)

before=PACKET/'measurements/baseline';after=PACKET/'measurements/candidate'
for path in before.glob('diff-*.tsv'):
    assert path.read_bytes()==(after/path.name).read_bytes(),path.name
    assert 'agree=yes' in path.read_text(),path.name
    if 'late-refusal' in path.name or 'malformed-escape' in path.name: assert 'refused:0:' in path.read_text(),path.name
for path in before.glob('scan-*.txt'): assert path.read_bytes()==(after/path.name).read_bytes(),path.name
results=[]
metrics=['allocation_calls','allocated_bytes','peak_live_delta','retained_live_delta']
for path in sorted(before.glob('alloc-*-0.tsv')):
    group=path.name[len('alloc-'):-len('-0.tsv')]
    allocations={}
    for label,folder in [('before',before),('after',after)]:
        rows=[row((folder/f'alloc-{group}-{i}.tsv').read_text()) for i in range(3)]
        for key in metrics+['callbacks','result','operation']:
            assert len({r[key] for r in rows})==1,(label,group,key)
        allocations[label]={k:int(rows[0][k]) for k in metrics}
        allocations[label]['semantics']=(rows[0]['result'],rows[0]['callbacks'])
    assert allocations['before']['semantics']==allocations['after']['semantics'],group
    timing={}
    for label,folder in [('aa',before),('abba',after)]:
        timing[label]={}
        for leg in (['a1','a2'] if label=='aa' else ['a1','b1','b2','a2']):
            summary,semantics=stats(folder/f'{leg}-{group}.tsv')
            assert semantics==allocations['before']['semantics'],(group,label,leg,semantics)
            timing[label][leg]=summary
    aa=timing['aa']; abba=timing['abba']
    percentages={metric:dict(b1_a1=delta(abba['b1'][metric],abba['a1'][metric]),b2_a2=delta(abba['b2'][metric],abba['a2'][metric]),aa=delta(aa['a2'][metric],aa['a1'][metric]),abba_a=delta(abba['a2'][metric],abba['a1'][metric]),abba_b=delta(abba['b2'][metric],abba['b1'][metric])) for metric in ['p50','mean','p95','p99']}
    results.append(dict(group=group,allocation=allocations,timing=timing,percent=percentages,review_regression=all(percentages['p50'][leg]>5 for leg in ['b1_a1','b2_a2'])))
assert len(results)==56,len(results)
layout={label:row((folder/'layout.tsv').read_text()) for label,folder in [('before',before),('after',after)]}
output=dict(layout=layout,groups=len(results),timing_samples=56*20*6,allocation_samples=56*3*2,results=results)
(PACKET/'comparison.json').write_text(json.dumps(output,indent=2)+'\n')
print('Verified source, constraints, probe, raw bindings, 10 differential cases, 56 workload groups; wrote comparison.json')
for result in results:
    if result['review_regression']: print('REVIEW >5% paired p50 regression:',result['group'],result['percent']['p50'])

# Native diagnostics bind the same before/after binaries and corpus as the matrix.
diagnostic=PACKET/'diagnostics'
for name,expected in json.loads((diagnostic/'raw-sha256.json').read_text()).items():
    assert digest(diagnostic/name)==expected,('diagnostic',name)
commands=json.loads((diagnostic/'commands.json').read_text())
assert len(commands)==4
for command,leg,phase in zip(commands,['a1','b1','b2','a2'],['baseline','candidate','candidate','baseline']):
    assert command['exit_code']==0
    assert command['binary_sha256']==manifests[phase]['binary_sha256']
    assert command['corpus_sha256']==manifests[phase]['corpus']['long-inline']['sha256']
    rows=[row(line) for line in (diagnostic/(leg+'.tsv')).read_text().splitlines()]
    assert len(rows)==1000
    assert [int(r['sample']) for r in rows]==list(range(1000))
    assert {(r['result'],r['callbacks']) for r in rows}=={('ok','256')}
print('Verified native diagnostic binary/corpus/raw bindings and 4,000 matching samples')
