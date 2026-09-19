#!/usr/bin/env python3
"""Verify frozen build, raw result, live candidate and retained diagnostic bindings."""
import hashlib
import json
import subprocess
from pathlib import Path
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
def sha(path): return hashlib.sha256(path.read_bytes()).hexdigest()
def read(name): return json.loads((P/name).read_text())
base=read('baseline.json')
for name,digest in base['constraints_sha256'].items(): assert sha(ROOT/name)==digest,name
for name,digest in base['build_inputs_sha256'].items():
    if (ROOT/name).exists(): assert sha(ROOT/name)==digest,name
builds={}
for phase in ['baseline','candidate']:
    rows=read(f'builds-{phase}.json')
    assert {x['label'] for x in rows}=={'native','allocations'}
    for row in rows:
        assert row['exit_code']==0
        builds[phase,row['label']]=row
        if Path(row['binary']).exists(): assert sha(Path(row['binary']))==row['binary_sha256']
        for name,digest in row['probe_sha256'].items(): assert sha(P/name)==digest,name
        if phase=='baseline': assert row['source_sha256']==base['source_sha256']
        else:
            for name,digest in row['source_sha256'].items(): assert sha(ROOT/name)==digest,name
native=read('native-runs-baseline.json')+read('native-runs-compare.json')
assert len(native)==78 and len({(r['case'],r['leg']) for r in native})==78
for row in native:
    assert row['exit_code']==0
    assert row['binary_sha256']==builds[row['phase'],'native']['binary_sha256']
    out=P/row['output']
    assert sha(out)==row['output_sha256']
    assert sha(out.with_suffix('.stderr'))==row['stderr_sha256']
    source=Path(row['command'][5])
    if row['source_sha256'] and source.exists(): assert sha(source)==row['source_sha256']
    if row['case'].endswith('-control'): assert row['source_sha256']==read('control-manifest.json')['control_sha256']
for phase in ['baseline','candidate']:
    rows=read(f'allocation-runs-{phase}.json'); assert len(rows)==13
    for row in rows:
        assert row['exit_code']==0
        assert row['binary_sha256']==builds[phase,'allocations']['binary_sha256']
        assert sha(P/row['output'])==row['output_sha256']
    profile=read(f'profile/{phase}/binding.json')
    assert profile['binary_sha256']==builds[phase,'native']['binary_sha256']
    for name,digest in profile['files'].items(): assert sha(P/'profile'/phase/name)==digest,name
for row in read('integration/results.json'):
    assert row['exit_code']==0,row['name']
    for name,digest in row['source_sha256'].items(): assert sha(ROOT/name)==digest,name
assert len(read('integration/results.json'))==7
for row in read('quality-summary.json'):
    assert row['exit_code']==0
    assert sha(P/'integration'/(row['name']+'.log'))==row['log_sha256']
    assert row['test_totals']['failed']==0
for row in read('evidence/results.json'):
    assert row['exit_code']==0,row['name']
    for name,digest in row['source_sha256'].items(): assert sha(ROOT/name)==digest,name
    assert sha(P/'evidence'/(row['name']+'.log'))==row['log_sha256']
assert len(read('evidence/results.json'))==6
derived=['native-summary.json','native-comparisons.json','semantic-bindings.json','native-leg-metadata.json','allocation-summary.json','allocation-comparisons.json','trace-summary.json','tables.md','native-review-triggers.json','baseline-noise.json','profile-summary.json']
prior={name:(P/name).read_bytes() for name in derived}
for script in ['summarize.py','summarize-allocations.py','trace-summary.py','report-metrics.py']:
    args = [str(x) for x in sorted((P/'trace-runs').iterdir()) if x.is_dir()] if script=='trace-summary.py' else []
    result=subprocess.run(['python3',str(P/script),*args],cwd=ROOT,capture_output=True)
    assert result.returncode==0,result.stderr.decode()
for name,contents in prior.items(): assert (P/name).read_bytes()==contents,name
reuse=read('baseline-trace-reuse.json')
for name,digest in reuse['files'].items(): assert sha(ROOT/name)==digest,name
trace=read('trace-summary.json')
assert trace['status']=='pass' and not trace['skipped']
assert len(trace['runs'])==1
for leg in trace['runs'][0]['legs']:
    expected=31 if leg['case'] in ['real','control','generated'] else None
    assert expected is not None
    assert len(leg['intervals']) == (1 if leg['operation']=='capture' else 3)
    for interval in leg['intervals']:
        assert interval['codec_calls']==expected
        assert interval['error_calls']==0
        assert interval['owned_calls']==(31 if leg['case']=='real' else 0)
        if leg['case']=='real' and interval['role']!='route_commit':
            assert interval['input_bytes']==550141
            assert interval['output_bytes']==634121
        if leg['case']=='real' and interval['role']=='route_commit':
            assert interval['input_bytes']==549999 and interval['output_bytes']==633981
control=read('control-manifest.json')
assert sha(ROOT/control['source'])==control['source_sha256']
if (P/control['control']).exists(): assert sha(P/control['control'])==control['control_sha256']
assert len(control['members'])==103 and sum(x['replacements'] for x in control['members'])==43
print('PASS: constraints, frozen builds, candidate source, all native/allocation samples, semantic consistency, profiles, trace restoration and gates.')
