#!/usr/bin/env python3
"""Replay the 0509 matched evidence without requiring deleted binaries."""
import copy
import hashlib
import json
import re
import sys
from pathlib import Path
HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
sys.path.insert(0, str(REPO))
from tools.summarize_crud_baseline import _validate_elapsed
from tools.validate_perf_corpus_binding import validate_binding

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def load(path):
    return json.loads(path.read_text())

def rss(path):
    return int(re.search(r'Maximum resident set size \(kbytes\): (\d+)', path.read_text())[1])

def main():
    manifests = {stage: load(HERE / stage / 'source-manifest.json' if stage == 'before' else HERE / 'source-manifest.json') for stage in ['before','after']}
    assert manifests['before'] == load(HERE.parent / 'change-0508/source-manifest.json')
    for path,digest in load(HERE / 'verifier-sources.json').items():
        assert sha(REPO / path) == digest, path
    changed = sorted(p for p in manifests['before'].keys() | manifests['after'].keys() if manifests['before'].get(p) != manifests['after'].get(p))
    assert changed == ['crates/litchi-odt/src/elements/text.rs', 'crates/litchi-odt/tests/sequential_text.rs'], changed
    for path,digest in manifests['after'].items():
        assert sha(REPO / path) == digest, path
    for path,digest in load(HERE / 'adr-manifest.json')['files'].items():
        assert sha(REPO / path) == digest, path
    builds = {stage: load(HERE / stage / 'build-receipt.json' if stage == 'before' else HERE / 'build-receipt.json') for stage in ['before','after']}
    assert builds['before']['binary_sha256'] == 'ab874929448dfa3a4c17fd9072e702081b4e78b84cae39f3be120518a0c7f6f2'
    for stage in builds:
        build_dir = HERE / 'before' if stage == 'before' else HERE
        b = builds[stage]
        assert b['exit_code'] == 0 and b['source_unchanged']
        assert sha(build_dir / 'source-manifest.json') == b['source_manifest_sha256']
        assert sha(build_dir / 'build.log') == b['log_sha256']
    reports = {}
    for stage in ['before','after']:
        for repeat in ['r1','r2']:
            directory = HERE / stage
            receipt = load(directory / f'{repeat}-receipt.json')
            assert receipt['binary_sha256'] == builds[stage]['binary_sha256']
            assert receipt['source_manifest_sha256'] == builds[stage]['source_manifest_sha256']
            assert receipt['exit_code'] == 0 and receipt['current_source_unchanged']
            for name,digest in receipt['artifacts'].items():
                assert sha(directory / name) == digest
            report = load(directory / f'{repeat}-report.json')
            catalog = load(directory / f'{repeat}-catalog.json')
            validate_binding(report,catalog)
            assert report['binary_identity']['binary_sha256'] == builds[stage]['binary_sha256']
            assert report['configuration']['samples_per_case'] == 500
            assert report['configuration']['warmup_iterations_per_case'] == 10
            assert receipt['command'][:3] == ['taskset','-c','2']
            rows = report['results']
            assert len(rows) == 3
            assert {r['corpus']['shape'] for r in rows} == {'tiny','medium','large'}
            for row in rows:
                assert row['case'] == 'odt_semantic_text_to_sink'
                _validate_elapsed(row, 500, stage + repeat)
                assert row['sink']['retained_output_bytes'] == 0
            reports[stage,repeat] = {r['corpus']['shape']:r for r in rows}
    for stage in ['before','after']:
        for name,count in [('heaptrack',20),('callgrind',5)]:
            report = load(HERE / stage / f'{name}-report.json')
            assert report['binary_identity']['binary_sha256'] == builds[stage]['binary_sha256']
            validate_binding(report,load(HERE / stage / f'{name}-catalog.json'))
            assert len(report['results']) == 1
            row = report['results'][0]
            _validate_elapsed(row,count,stage+name)
            for field in ['corpus','sink','output_sha256']:
                assert row[field] == reports[stage,'r1']['large'][field]
    profile = load(HERE / 'profile-summary.json')
    for stage,count,ir in [('before',200000,304455496),('after',20,298891534)]:
        assert int(re.search(r'(\d+) calls to allocation functions with',(HERE/stage/'heaptrack-sink-buffer.txt').read_text())[1]) == count
        assert int(re.search(r'^summary: (\d+)',(HERE/stage/'callgrind.out').read_text(),re.M)[1]) == ir
        assert profile[stage]['append_sink_precharged_allocation_calls'] == count
        assert profile[stage]['callgrind_instruction_references'] == ir
    pairs = []
    for repeat in ['r1','r2']:
        for shape in ['tiny','medium','large']:
            before,after = [reports[stage,repeat][shape] for stage in ['before','after']]
            for field in ['corpus','sink','output_sha256']:
                assert before[field] == after[field], field
            values = {'repeat':repeat,'shape':shape,'before_p50_ns':before['elapsed_ns']['p50'],'after_p50_ns':after['elapsed_ns']['p50']}
            for metric in ['p50','p95','p99','mean']:
                values[metric+'_change_percent'] = (after['elapsed_ns'][metric]/before['elapsed_ns'][metric]-1)*100
            values['throughput_change_percent'] = (before['elapsed_ns']['mean']/after['elapsed_ns']['mean']-1)*100
            values['adverse_latency_or_throughput_over_5_percent'] = any(values[m+'_change_percent']>5 for m in ['p50','p95','p99','mean']) or values['throughput_change_percent'] < -5
            pairs.append(values)
    memory = []
    for repeat in ['r1','r2']:
        b,a = [rss(HERE / stage / f'{repeat}.log') for stage in ['before','after']]
        memory.append({'repeat':repeat,'scope':'whole child, all three shapes and setup','before_peak_rss_kib':b,'after_peak_rss_kib':a,'change_percent':(a/b-1)*100,'adverse_over_5_percent':a/b>1.05})
    bad = copy.deepcopy(reports['after','r1']['large'])
    bad['elapsed_ns']['samples'].pop()
    try:
        _validate_elapsed(bad,500,'negative short vector')
    except (ValueError,AssertionError):
        pass
    else:
        raise AssertionError('short sample vector accepted')
    drift = []
    for stage in ['before','after']:
        for shape in ['tiny','medium','large']:
            a,b = [reports[stage,r][shape]['elapsed_ns'] for r in ['r1','r2']]
            changes = {m:(b[m]/a[m]-1)*100 for m in ['p50','mean','p95','p99']}
            drift.append({'stage':stage,'shape':shape,'repeat_change_percent':changes,'exceeds_drift_ceiling':any(abs(changes[m])>limit for m,limit in [('p50',5),('mean',5),('p95',10),('p99',15)])})
    for name in ['fmt','odt-tests','odt-clippy','odt-rustdoc','harness-tests','harness-clippy']:
        gate = load(HERE / f'{name}-receipt.json')
        assert gate['exit_code'] == 0 and gate['source_unchanged']
        assert gate['source_manifest_sha256'] == builds['after']['source_manifest_sha256']
        assert gate['log_sha256'] == sha(HERE / f'{name}.log')
    gates = load(HERE / 'gates.json')
    assert gates['total_passed'] == 1491 and gates['total_ignored'] == 1
    for name,receipt in gates['source_bound_receipts'].items():
        assert receipt == load(HERE / f'{name}-receipt.json')
    for name,receipt in gates['additional_checks'].items():
        assert receipt['exit_code'] == 0 and receipt['log_sha256'] == sha(HERE / name)
    for name,expected in [('odt-tests',1008),('harness-tests',483)]:
        counts = re.findall(r'test result: ok\. (\d+) passed; (\d+) failed; (\d+) ignored;', (HERE/f'{name}.log').read_text())
        assert sum(int(c[0]) for c in counts) == expected
        assert sum(int(c[1]) for c in counts) == 0
    tail = {}
    tail_pairs = []
    for stage in ['before','after']:
        for repeat in ['tail-r1','tail-r2']:
            directory = HERE / stage
            receipt = load(directory / f'{repeat}-receipt.json')
            assert receipt['binary_sha256'] == builds[stage]['binary_sha256']
            assert receipt['source_manifest_sha256'] == builds[stage]['source_manifest_sha256']
            assert receipt['exit_code'] == 0 and receipt['current_source_unchanged']
            assert receipt['command'][:3] == ['taskset','-c','2']
            for name,digest in receipt['artifacts'].items():
                assert sha(directory/name) == digest
            report = load(directory / f'{repeat}-report.json')
            validate_binding(report,load(directory / f'{repeat}-catalog.json'))
            assert report['binary_identity']['binary_sha256'] == builds[stage]['binary_sha256']
            assert report['configuration']['samples_per_case'] == 2000
            assert report['configuration']['warmup_iterations_per_case'] == 10
            assert len(report['results']) == 1
            row = report['results'][0]
            _validate_elapsed(row,2000,stage+repeat)
            for field in ['corpus','sink','output_sha256']:
                assert row[field] == reports[stage,'r1']['large'][field]
            tail[stage,repeat] = row['elapsed_ns']
    for repeat in ['tail-r1','tail-r2']:
        b,a = [tail[stage,repeat] for stage in ['before','after']]
        br,ar = [rss(HERE / stage / f'{repeat}.log') for stage in ['before','after']]
        changes = {m:(a[m]/b[m]-1)*100 for m in ['p50','mean','p95','p99']}
        changes['throughput'] = (b['mean']/a['mean']-1)*100
        changes['whole_child_peak_rss'] = (ar/br-1)*100
        tail_pairs.append({'repeat':repeat,'changes_percent':changes,'before_p99_ns':b['p99'],'after_p99_ns':a['p99'],'before_peak_rss_kib':br,'after_peak_rss_kib':ar,'adverse_over_5_percent':any(changes[m]>5 for m in ['p50','mean','p95','p99','whole_child_peak_rss']) or changes['throughput'] < -5})
    tail_drift = []
    for stage in ['before','after']:
        a,b = [tail[stage,r] for r in ['tail-r1','tail-r2']]
        changes = {m:(b[m]/a[m]-1)*100 for m in ['p50','mean','p95','p99']}
        tail_drift.append({'stage':stage,'changes_percent':changes,'exceeds_drift_ceiling':any(abs(changes[m])>limit for m,limit in [('p50',5),('mean',5),('p95',10),('p99',15)])})
    assert sha(HERE / 'initial-summary.json') == load(HERE / 'tail-followup-admission.json')['initial_summary_sha256']
    print(json.dumps({'tail_followup_samples':8000,'tail_pairs':tail_pairs,'tail_drift':tail_drift,'drift':drift,'samples':6000,'paired_rows':6,'source_changes':changed,'pairs':pairs,'whole_child_rss':memory,'negative_short_vector_rejected':True},indent=2))

if __name__ == '__main__':
    main()
