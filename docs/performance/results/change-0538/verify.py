"""Read-only replay of planning allocation smoke and quality evidence."""
import hashlib
import json
from pathlib import Path
import re
import subprocess

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def verify():
    before = {str(p.relative_to(HERE)): sha(p) for p in HERE.rglob('*') if p.is_file()}
    for name, digest in json.loads((HERE/'source-manifest.json').read_text()).items():
        assert sha(REPO/name) == digest, name
    for name, digest in json.loads((HERE/'adr-manifest.json').read_text())['files'].items():
        assert sha(REPO/name) == digest, name
    plan = json.loads((HERE/'plan.json').read_text())
    assert not subprocess.check_output(['git','diff',plan['base_revision'],'--name-only','--','crates'], cwd=REPO)
    names = ['test-xlsx','test-xlsx-tmpdir','test-allocator','test-wrapper','check','clippy','rustdoc',
             'fmt','boundaries','build-smoke','test-integration']
    names += [binary+'-'+shape for binary in ['normal','alloc'] for shape in ['medium','dense-sparse']]
    assert {p.name for p in HERE.glob('*.receipt.json')} == {name+'.receipt.json' for name in names}
    receipts = [json.loads((HERE/(name+'.receipt.json')).read_text()) for name in names]
    assert receipts[0]['exit_code'] == 101
    assert all(r['exit_code'] == 0 for r in receipts[1:])
    assert all(r['finished'] >= r['started'] for r in receipts)
    initial = (HERE/'test-xlsx.log').read_text()
    assert 'QuotaExceeded' in initial and '42 passed; 1 failed;' in initial
    failure = json.loads((HERE/'initial-failure.json').read_text())
    assert failure['removed'] and not Path(failure['removed_empty_fixture']).exists()
    assert failure['failed_test'] in initial
    chronological = sorted(receipts,key=lambda r:r['started'])
    assert all(a['finished'] <= b['started'] for a,b in zip(chronological,chronological[1:]))
    source_sha = sha(REPO/'tools/perf-baseline/src/lib.rs')
    assert all(r['source_sha256'] == source_sha for r in receipts)
    test_counts = {}
    for name in names:
        if name.startswith('test-'):
            if name == 'test-xlsx':
                test_counts[name] = 42
                continue
            counts = re.findall(r'test result: ok\. (\d+) passed; 0 failed;', (HERE/(name+'.log')).read_text())
            assert counts, name
            test_counts[name] = sum(map(int,counts))
    expected_cases = {'xlsx_source_backed_'+managed+'cell_values_'+edit+'_edit_save'
                      for managed in ['', 'managed_'] for edit in ['one','one_percent','batch','multi_sheet']}
    numeric = {'allocation_calls','deallocation_calls','reallocation_calls','failed_allocation_calls',
               'allocated_bytes','deallocated_bytes','live_bytes_before','live_bytes_after',
               'peak_live_bytes_before','peak_live_bytes_after','region_peak_live_bytes'}
    rows = []
    identities = {}
    binaries = json.loads((HERE/'binary-manifest.json').read_text())
    for binary in ['normal','alloc']:
        for shape in ['medium','dense-sparse']:
            report = json.loads((HERE/(binary+'-'+shape+'.json')).read_text())
            assert report['configuration']['warmup_iterations_per_case'] == 1
            assert report['configuration']['samples_per_case'] == 3
            assert report['configuration']['xlsx_cell_crud_shapes'] == [shape]
            assert report['tool']['profile'] == 'debug'
            assert report['tool']['binary'] == 'litchi-perf-baseline'+('-alloc' if binary == 'alloc' else '')
            assert report['tool']['instrumentation'] == ('system_allocator_operation_scoped' if binary == 'alloc' else 'none')
            executable = binaries[report['tool']['binary']]
            assert report['binary_identity']['binary_sha256'] == executable['sha256']
            assert report['binary_identity']['binary_bytes'] == executable['bytes']
            assert report['binary_identity']['path'] == executable['path']
            results = report['results']
            assert len(results) == 8 and {r['case'] for r in results} == expected_cases
            for result in results:
                case = result['case']
                e = result['source']['xlsx_cell_values']
                timing = result['elapsed_ns']
                assert len(timing['samples']) == 3 and sorted(timing['sample_order']) == [0,1,2]
                for phase in ['open','plan','commit','publication','reopen']:
                    assert len(e[phase+'_ns']) == 3
                for sorted_index, acquisition in enumerate(timing['sample_order']):
                    assert timing['samples'][sorted_index] == sum(e[p+'_ns'][acquisition] for p in ['open','plan','commit','publication'])
                for phase in ['plan','commit','publication']:
                    metrics = e[phase+'_allocation_metrics']
                    assert len(metrics) == 3
                    for sample in metrics:
                        assert sample['scope'] == 'operation_global_system_allocator'
                        if binary == 'normal':
                            assert sample == dict(status='unavailable',scope='operation_global_system_allocator')
                        else:
                            assert sample['status'] == 'measured'
                            assert set(sample) == numeric | {'scope','status'}
                            assert all(type(sample[k]) is int and sample[k] >= 0 for k in numeric)
                            assert sample['allocation_calls'] > 0 and sample['allocated_bytes'] > 0
                            assert sample['failed_allocation_calls'] == 0
                            assert sample['region_peak_live_bytes'] >= max(sample['live_bytes_before'],sample['live_bytes_after'])
                            assert sample['peak_live_bytes_after'] >= sample['peak_live_bytes_before']
                            assert sample['peak_live_bytes_after'] >= sample['region_peak_live_bytes']
                            assert sample['live_bytes_before'] + sample['allocated_bytes'] - sample['deallocated_bytes'] == sample['live_bytes_after']
                assert e['budget_used_after_handles_drop'] == [0]*3
                assert e['budget_objects_used_after_handles_drop'] == [0]*3
                identity = dict(corpus=result['corpus'],output=result['output_sha256'],
                    semantic=e['semantic_sha256'],untouched=e['untouched_member_sha256'],
                    source_reads=e['source_read_bytes'],selected=e['selected_worksheet_count'])
                key = (shape,case)
                if key in identities:
                    assert identities[key] == identity, key
                else:
                    identities[key] = identity
                rows.append(dict(binary=binary,shape=shape,case=case,samples=3,
                    planning_allocation_status=e['plan_allocation_metrics'][0]['status']))
    if (HERE/'cleanup.json').exists():
        cleanup = json.loads((HERE/'cleanup.json').read_text())
        assert cleanup['removed'] and not Path(cleanup['path']).exists()
    if (HERE/'SHA256SUMS').exists():
        assert (HERE/'cleanup.json').exists(), 'sealed batch must record cleanup'
        seal = dict(line.split('  ',1)[::-1] for line in (HERE/'SHA256SUMS').read_text().splitlines())
        assert seal == {k:v for k,v in before.items() if k != 'SHA256SUMS'}
    assert before == {str(p.relative_to(HERE)): sha(p) for p in HERE.rglob('*') if p.is_file()}
    return dict(status='pass',scope='Debug correctness/schema smoke only; no performance comparison',
                receipts=len(receipts),successful_receipts=len(receipts)-1,
                initial_environment_failure='Default /tmp quota; exact test passed with owned TMPDIR',
                test_counts=test_counts,records=len(rows),
                measured_planning_samples=48,unavailable_planning_samples=48,
                runtime_changed=False,source_manifest_sha256=sha(HERE/'source-manifest.json'))

if __name__ == '__main__':
    print(json.dumps(verify(),indent=2,sort_keys=True))
