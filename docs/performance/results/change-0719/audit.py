#!/usr/bin/env python3
"""Replay source, command/log custody and producer output identity checks."""
import importlib.util,json,sys
from pathlib import Path
P=Path(__file__).resolve().parent
s=importlib.util.spec_from_file_location('custody0719audit',P/'custody.py');C=importlib.util.module_from_spec(s);s.loader.exec_module(C)
def read(name):return json.loads((P/name).read_text())
def analyze():
    source=read('source.json');assert source==C.census()
    before=read('baseline-source.json');assert set(source)==set(before)
    changed=[n for n in source if source[n]!=before[n]]
    assert changed==['tools/perf-baseline/src/producer_shape.rs'],changed
    for manifest in ['constraints.json','motivation.json']:
        for n,h in read(manifest).items():assert C.sha(C.ROOT/n)==h,n
    assert read('review.json')['final_review']['source_sha256']==source['tools/perf-baseline/src/producer_shape.rs']
    for folder in ['initial-oracle-build','initial-workbook-change-assumption','initial-quality-warning']:
        archived=read(folder+'/source.json')
        assert C.sha(P/folder/'producer_shape.rs.txt')==archived['tools/perf-baseline/src/producer_shape.rs']
        for receipt in (P/folder).glob('*.receipt.json'):
            r=json.loads(receipt.read_text());name=receipt.name.removesuffix('.receipt.json')
            assert r['source_sha256']==C.sha(P/folder/'source.json')
            assert r['log_sha256']==C.sha(P/folder/(name+'.log'))
    for n,h in read('documentation.json').items():assert C.sha(C.ROOT/n)==h,n
    gate=read('final-report-gate.json');assert gate['exit_code']==0
    assert gate['log_sha256']==C.sha(P/'final-report-gate.log')
    assert gate['documentation_sha256']==C.sha(P/'documentation.json')
    binary=read('binary.json');path=Path(binary['path'])
    if path.exists():assert C.sha(path)==binary['sha256'] and path.stat().st_size==binary['bytes']
    else:
        cleanup=read('cleanup.json');assert cleanup['binary']==binary
        assert all(not Path(n).exists() for n in cleanup['removed_paths'])
    receipts={p.name:json.loads(p.read_text()) for p in P.glob('*.receipt.json')}
    names=['build','fmt','tests','clippy','rustdoc']
    names += ['gate-'+r['name'] for r in json.loads((P.parent/'change-0717/evidence/results.json').read_text())]
    names += [f'correctness-r{r}-{shape}' for r in [1,2] for shape in ['medium','dense']]
    assert set(receipts)=={n+'.receipt.json' for n in names}
    for name in names:
        row=receipts[name+'.receipt.json'];assert row['exit_code']==0,name
        assert row['source_sha256']==C.sha(P/'source.json') and row['runner_sha256']==C.sha(P/'run.py')
        assert row['log_sha256']==C.sha(P/(name+'.log'))
    prior=read('prior-producer.json');assert C.sha(C.ROOT/prior['path'])==prior['sha256']
    historical=json.loads((C.ROOT/prior['path']).read_text())['results'][0]
    rows=[]
    for shape in ['medium','dense']:
        same=[]
        for r in [1,2]:
            name=f'correctness-r{r}-{shape}';p=read(name+'.json')
            assert p['tool']['instrumentation']=='none'
            assert p['binary_identity']['binary_sha256']==binary['sha256']
            assert p['configuration']['samples_per_case']==3 and p['configuration']['warmup_iterations_per_case']==1
            assert len(p['results'])==1
            result=p['results'][0]
            assert result['case']==f'xlsx_producer_{shape}_source_one_edit_save'
            expected=[binary['path'],'--warmup','1','--samples','3','--case',result['case'],'--producer-evidence',str(P/(name+'.corpus.json')),'--json',str(P/(name+'.json'))]
            assert receipts[name+'.receipt.json']['command']==expected
            assert len(result['elapsed_ns']['samples'])==3
            identity={k:result[k] for k in ['case','corpus','sink','output_sha256']}
            if shape=='medium':
                assert identity['corpus']==historical['corpus'] and identity['output_sha256']==historical['output_sha256']
            assert len(identity['output_sha256'])==64 and identity['sink']['accepted_bytes']>0
            read(name+'.corpus.json')
            same.append(identity)
            rows.append(dict(name=name,result_sha256=C.sha(P/(name+'.json')),corpus_evidence_sha256=C.sha(P/(name+'.corpus.json')),identity=identity))
        assert same[0]==same[1],shape
    return dict(status='pass',changed_source=changed,children=4,retained_samples=12,warmups=4,performance_claim='none',results=rows)
if __name__=='__main__':
    value=analyze()
    if sys.argv[1:]==['--write']:(P/'analysis.json').write_text(json.dumps(value,indent=2)+'\n')
    else:assert value==read('analysis.json')
    print('PASS: source, command/log custody, four producer children and output parity')
