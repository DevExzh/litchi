"""Replay the rejected pilot and final source/quality custody without mutation."""
import hashlib
import json
from pathlib import Path
import re
import subprocess
import analyze as native
import analyze_allocation
import run

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def read(name):
    return json.loads((HERE/name).read_text())

def verify():
    before = {str(p.relative_to(HERE)):sha(p) for p in HERE.rglob('*') if p.is_file()}
    comparison = native.analyze()
    assert comparison == read('comparison.json')
    allocation = analyze_allocation.analyze()
    assert allocation == read('allocation-analysis.json')
    assert comparison['status'] == 'pass' and not comparison['native_admission']['passed']
    assert allocation['planning_gate_passed']
    assert sum(r['passed'] for r in comparison['native_admission']['rows']) == 1
    decision = read('decision.json')
    assert decision['decision'] == 'reject' and decision['runtime_restored']
    assert decision['comparison_sha256'] == sha(HERE/'comparison.json')
    assert decision['allocation_analysis_sha256'] == sha(HERE/'allocation-analysis.json')
    review = read('adverse-review.json')
    assert review['status'] == 'complete' and review['disposition'] == 'reject'
    for field, filename in [('comparison_sha256','comparison.json'), ('allocation_analysis_sha256','allocation-analysis.json'), ('plan_sha256','plan.json')]:
        assert review[field] == sha(HERE/filename)
    for field, source in [('adverse_flags','adverse_flags_over_five_percent'), ('same_build_drift_flags','same_build_drift_over_five_percent')]:
        rows = review[field]
        originals = [row['original'] for row in rows]
        key = lambda value: json.dumps(value, sort_keys=True)
        assert sorted(originals, key=key) == sorted(comparison['comparison'][source], key=key)
        assert len({row['id'] for row in rows}) == len(rows)
        assert all(row['interpretation'].strip() for row in rows)
    run.check_source('baseline')
    baseline = read('baseline/source-manifest.json')
    candidate = read('candidate/source-manifest.json')
    assert (HERE/'baseline/source.patch').read_bytes() == (HERE/'guard.patch').read_bytes()
    assert run.candidate_files(baseline,candidate) == {'crates/litchi-xlsx/src/raw/worksheet/codec.rs'}
    codec = REPO/'crates/litchi-xlsx/src/raw/worksheet/codec.rs'
    assert sha(codec) == decision['codec_sha256'] == baseline[str(codec.relative_to(REPO))]
    plan = read('plan.json')
    assert codec.read_bytes() == subprocess.check_output(['git','show',plan['revision']+':'+str(codec.relative_to(REPO))],cwd=REPO)
    for name,digest in read('adr-manifest.json')['files'].items():
        assert sha(REPO/name) == digest,name
    frozen = read('frozen-inputs.json')
    for name,digest in frozen['files'].items():
        assert sha(HERE/name) == digest,name
    jobs = native._expected_native_jobs(plan)
    alloc_jobs = native._expected_allocation_jobs(plan)
    quality = read('quality-plan.json')['commands']
    receipts = []
    test_counts = {}
    for stage in ['baseline','candidate']:
        expected = {j['name'] for j in jobs+alloc_jobs} | {'build-normal','build-alloc','quality-tests'}
        if stage == 'baseline':
            expected |= {'final-'+name for name in quality}
        actual = {p.name.removesuffix('.receipt.json') for p in (HERE/stage).glob('*.receipt.json')}
        assert actual == expected,(stage,actual^expected)
        for name in sorted(expected):
            r = read(stage+'/'+name+'.receipt.json')
            assert r['exit_code'] == 0
            assert r['source_manifest_sha256'] == sha(HERE/stage/'source-manifest.json')
            working = 'candidate' if stage == 'candidate' or (name.startswith(('native-r2-','alloc-r2-')) and stage == 'baseline') else 'baseline'
            assert r['working_source_manifest_sha256'] == sha(HERE/working/'source-manifest.json')
            assert r['script_sha256'] == sha(HERE/'run.py') and r['plan_sha256'] == sha(HERE/'plan.json')
            for artifact,digest in r['artifacts'].items():
                assert sha(HERE/stage/artifact) == digest
            if name == 'quality-tests' or name.startswith('final-'):
                key = name.removeprefix('final-')
                assert r['command'] == quality[key]
                if key == 'quality-tests':
                    text = (HERE/stage/(name+'.stdout')).read_text()
                    count = sum(map(int,re.findall(r'test result: ok\. (\d+) passed;',text)))
                    assert count == 1292
                    test_counts[stage+'/'+name] = count
            receipts.append(r)
        for job in jobs:
            report = read(stage+'/'+job['name']+'.json')
            source = report['results'][0]['source']['xlsx_cell_values']
            for phase in ['plan','commit','publication']:
                samples = source[phase+'_allocation_metrics']
                assert len(samples) == len(source[phase+'_ns']) == job['samples']
                assert all(s == dict(status='unavailable',scope='operation_global_system_allocator') for s in samples)
    ordered = sorted(receipts,key=lambda r:r['start_utc'])
    assert all(a['end_utc'] <= b['start_utc'] for a,b in zip(ordered,ordered[1:]))
    captures = [r for r in ordered if r['binary_sha256'] is not None]
    for lane in ['native','alloc']:
        previous_end = ''
        for stage,repeat in [('baseline',1),('candidate',1),('candidate',2),('baseline',2)]:
            group = [read(str(p.relative_to(HERE))) for p in (HERE/stage).glob(f'{lane}-r{repeat}-*.receipt.json')]
            assert group and previous_end < min(r['start_utc'] for r in group)
            previous_end = max(r['end_utc'] for r in group)
    assert frozen['created_utc'] < min(r['start_utc'] for r in captures)
    assert read('allocation-gates.json')['created_utc'] < min(r['start_utc'] for r in captures)
    # Source restoration is after every timed child, before final checks.
    final_receipts = [read('baseline/final-'+name+'.receipt.json') for name in quality]
    assert max(r['end_utc'] for r in captures) < min(r['start_utc'] for r in final_receipts)
    assert not list(HERE.glob('*/profile-*.receipt.json'))
    assert not list(HERE.glob('*/hardware-*.receipt.json'))
    assert not list(HERE.glob('*/eager-*.receipt.json'))
    if (HERE/'cleanup.json').exists():
        cleanup = read('cleanup.json')
        assert cleanup['removed'] == plan['owned_paths']
        assert cleanup['owned_paths_absent'] and cleanup['accessible_process_references'] == []
        assert all(not Path(p).exists() for p in plan['owned_paths'])
    if (HERE/'SHA256SUMS').exists():
        assert (HERE/'cleanup.json').exists()
        sealed = dict(line.split('  ',1)[::-1] for line in (HERE/'SHA256SUMS').read_text().splitlines())
        assert sealed == {k:v for k,v in before.items() if k != 'SHA256SUMS'}
    assert before == {str(p.relative_to(HERE)):sha(p) for p in HERE.rglob('*') if p.is_file()}
    return dict(status='pass',decision='reject',receipts=len(receipts),native_samples=2440,
        allocator_samples_per_phase=240,test_counts=test_counts,total_tests=sum(test_counts.values()),
        matched_adverse=len(comparison['comparison']['adverse_flags_over_five_percent']),
        same_build_drift=len(comparison['comparison']['same_build_drift_over_five_percent']),
        runtime_restored=True,baseline_manifest_sha256=sha(HERE/'baseline/source-manifest.json'))

if __name__ == '__main__':
    print(json.dumps(verify(),indent=2,sort_keys=True))
