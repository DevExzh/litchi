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
for phase in ['baseline','candidate']:
    row=read(f'build-refusal-{phase}.json')
    assert row['exit_code']==0
    assert row['source_sha256']==builds[phase,'native']['source_sha256']
    for name,digest in row['probe_sha256'].items(): assert sha(P/name)==digest,name
    if Path(row['binary']).exists(): assert sha(Path(row['binary']))==row['binary_sha256']
refusal=read('refusal-runs-baseline.json')+read('refusal-runs-compare.json')
assert len(refusal)==6 and len({r['leg'] for r in refusal})==6
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
derived=['refusal-tables.md','refusal-summary.json','refusal-bindings.json','refusal-comparisons.json','refusal-review-triggers.json','native-summary.json','native-comparisons.json','semantic-bindings.json','native-leg-metadata.json','allocation-summary.json','allocation-comparisons.json','trace-summary.json','tables.md','native-review-triggers.json','baseline-noise.json','profile-summary.json']
prior={name:(P/name).read_bytes() for name in derived}
for script in ['summarize.py','summarize-allocations.py','summarize-refusal.py','trace-summary.py','report-metrics.py']:
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
    expected=19 if leg['case']=='generated' else 18 if leg['case'] in ['real','control'] else None
    assert expected is not None
    assert len(leg['intervals']) == (1 if leg['operation']=='capture' else 3)
    for interval in leg['intervals']:
        assert interval['codec_calls']==expected
        assert interval['error_calls']==0
        assert interval['owned_calls']==(18 if leg['case']=='real' else 0)
        assert interval['layout'] and len(interval['layout'])==1
        assert interval['layout'][0]['proof_entry_bytes']==24
        assert interval['layout'][0]['slide_entry_bytes']==48
        if leg['case']=='real':
            assert interval['input_bytes']==(280892 if interval['role']=='route_commit' else 280963)
            assert interval['output_bytes']==(323993 if interval['role']=='route_commit' else 324063)
control=read('control-manifest.json')
assert sha(ROOT/control['source'])==control['source_sha256']
if (P/control['control']).exists(): assert sha(P/control['control'])==control['control_sha256']
assert len(control['members'])==103 and sum(x['replacements'] for x in control['members'])==43
# Superseded helper measurements remain independently bound to their archived source.
initial=P/'initial-identity-helper'
initial_builds={r['label']:r for r in json.loads((initial/'builds-candidate.json').read_text())}
for row in initial_builds.values():
    for name,digest in row['source_sha256'].items():
        if digest!=base['source_sha256'][name]: assert sha(initial/'sources'/name)==digest,name
    for name,digest in row['probe_sha256'].items(): assert sha(P/name)==digest,name
for prefix in ['native','allocation','refusal']:
    manifests = ['baseline','compare'] if prefix!='allocation' else ['baseline','candidate']
    for suffix in manifests:
        for row in json.loads((initial/f'{prefix}-runs-{suffix}.json').read_text()):
            assert row['exit_code']==0
            path=initial/row['output']
            assert sha(path)==row['output_sha256']
            if 'stderr_sha256' in row: assert sha(path.with_suffix('.stderr'))==row['stderr_sha256']
            if prefix=='refusal':
                binding=json.loads((initial/'build-refusal-candidate.json').read_text()) if row['phase']=='candidate' else read('build-refusal-baseline.json')
            else:
                label='native' if prefix=='native' else 'allocations'
                phase=row['phase']
                binding=initial_builds[label] if phase=='candidate' else builds[phase,label]
            assert row['binary_sha256']==binding['binary_sha256']
for phase in ['baseline','candidate']:
    binding=json.loads((initial/'profile'/phase/'binding.json').read_text())
    for name,digest in binding['files'].items(): assert sha(initial/'profile'/phase/name)==digest,name
expanded=read('expanded-refusal-baseline/manifest.json')
assert expanded['status']=='built' and expanded['restored_exact']
for row in expanded['sources']:
    stem=row['path'].replace('/','__')
    for phase in ['baseline','candidate']:
        assert sha(P/'expanded-refusal-baseline'/(stem+'.'+phase))==row[phase+'_sha256']
    assert row['baseline_sha256']==base['source_sha256'][row['path']]
for path in (P/'trace-initial').glob('*/manifest.json'):
    attempt=json.loads(path.read_text())
    assert attempt['status']=='failed' and attempt['source_restored_exact'] and not attempt['probe_runs']
print('PASS: constraints, frozen builds, current and superseded candidate sources, native/allocation/refusal samples, profiles, trace restoration and gates.')
