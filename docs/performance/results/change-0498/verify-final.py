#!/usr/bin/env python3
"""Independently check final child receipts, byte oracles, counters and hashes."""
import csv
import hashlib
import json
from pathlib import Path

HERE=Path(__file__).resolve().parent
ROOT=HERE.parents[3]
freeze=json.loads((HERE/'after-freeze.json').read_text())
assert hashlib.sha256(Path(freeze['binary']).read_bytes()).hexdigest()==freeze['binary_sha256']
for item in json.loads((HERE/'candidate-source.json').read_text())['files']:
    assert hashlib.sha256((ROOT/item['path']).read_bytes()).hexdigest()==item['sha256'],item['path']
children=[]
for corpus in ['few-large','many-small']:
    for source in ['owned','file','instrumented']:
        reference=None
        for api,workers in [('serial',1),('batch',1),('batch',2),('batch',4),('batch',8)]:
            name=f'{corpus}-{source}-{api}-w{workers}'
            p=HERE/'final'/name
            receipt=json.loads(p.with_suffix('.json').read_text())
            assert receipt['exit_code']==0 and receipt['cleanup_verified'],name
            assert receipt['binary_sha256']==freeze['binary_sha256'],name
            assert hashlib.sha256(p.with_suffix('.csv').read_bytes()).hexdigest()==receipt['csv_sha256'],name
            assert hashlib.sha256(p.with_suffix('.time.stderr').read_bytes()).hexdigest()==receipt['stderr_sha256'],name
            rows=list(csv.DictReader(p.with_suffix('.csv').open()))
            samples=[r for r in rows if r['record']=='sample']
            assert len(samples)==66,name
            assert {(int(r['repeat']),int(r['index'])) for r in samples}=={(a,b) for a in [0,1] for b in range(33)},name
            keys=['corpus_sha256','logical_bytes','digest','source_calls','source_requested_bytes',
                  'source_returned_bytes','cache_cold_loads','cache_hits','budget_input_bytes','budget_work']
            signatures={tuple(r[k] for k in keys) for r in samples}
            assert len(signatures)==1,name
            if reference is None:reference=signatures
            assert signatures==reference,name
            assert all(int(r['budget_memory_after_release'])==0 and int(r['budget_objects_after_release'])==0 for r in samples),name
            children.append(name)
assert len(list((HERE/'final').glob('*.json')))==len(children)==30
result={'status':'pass','children':len(children),'measured_samples':len(children)*60,
        'warmup_samples':len(children)*6,'source_counter_and_budget_work_equality':True,
        'byte_oracles_and_released_budgets':True,'candidate_source_and_raw_hashes':True}
(HERE/'final-verification.json').write_text(json.dumps(result,indent=2)+'\n')
print(json.dumps(result))
