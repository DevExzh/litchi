"""Validate current native/allocation CFB/OLE2 evidence without speedup claims."""
import argparse
import datetime
import importlib.util
import json
from pathlib import Path

from run import HERE, FOLDER, REPO, SCRATCH, TARGET, jobs, sha

HELPER = HERE.parent / 'change-0511/verify.py'
spec = importlib.util.spec_from_file_location('cfb_0511_checks', HELPER)
OLD = importlib.util.module_from_spec(spec)
spec.loader.exec_module(OLD)
METRICS = ('p50', 'p95', 'p99', 'mean')


def read(path):
    return json.loads(path.read_text())


def binary(kind):
    meta = read(FOLDER / ('binary-' + kind + '.json'))
    build_path = FOLDER / ('build-' + kind + '.receipt.json')
    assert meta['build_receipt_sha256'] == sha(build_path)
    assert meta['source_manifest_sha256'] == sha(FOLDER / 'source-manifest.json')
    expected = ['env', 'CARGO_BUILD_JOBS=2', 'CARGO_INCREMENTAL=0', 'cargo', 'build',
        '--release', '--locked', '--manifest-path', 'tools/perf-baseline/Cargo.toml',
        '--bin', 'litchi-perf-baseline' + ('-alloc' if kind == 'alloc' else ''),
        '--target-dir', str(TARGET)]
    if kind == 'alloc':
        expected += ['--features', 'allocator-metrics']
    assert read(build_path)['command'] == expected
    receipt('build-' + kind)
    path = Path(meta['path'])
    assert path == SCRATCH / kind
    if path.exists():
        assert sha(path) == meta['sha256'] and path.stat().st_size == meta['bytes']
    else:
        cleanup = read(HERE / 'cleanup.json')
        assert cleanup['plan_sha256'] == sha(HERE / 'plan.json')
        assert cleanup['owned_paths_absent'] and cleanup['removed'] == read(HERE / 'plan.json')['owned_paths']
        assert all(not __import__('os').path.lexists(p) for p in cleanup['removed'])
    return meta


def receipt(name, kind=None, allow_failure=False):
    path = FOLDER / (name + '.receipt.json')
    r = read(path)
    assert allow_failure or r['exit_code'] == 0, name
    assert r['script_sha256'] == sha(HERE / 'run.py')
    assert r['plan_sha256'] == sha(HERE / 'plan.json')
    assert r['source_manifest_sha256'] == sha(FOLDER / 'source-manifest.json')
    assert r['seconds'] > 0
    assert datetime.datetime.fromisoformat(r['start_utc']) < datetime.datetime.fromisoformat(r['end_utc'])
    assert all(v is None for v in r['environment'].values()), r['environment']
    if kind:
        assert r['binary_sha256'] == read(FOLDER / ('binary-' + kind + '.json'))['sha256']
    else:
        assert r['binary_sha256'] is None
    for n, digest in r['artifacts'].items():
        assert Path(n).name == n and sha(FOLDER / n) == digest, n
    return r


def normalized(value):
    if isinstance(value, list):
        assert value and all(v == value[0] for v in value), 'source vector varies'
        return normalized(value[0])
    if isinstance(value, dict):
        return {k: normalized(v) for k, v in value.items()}
    return value


def identity(row):
    return dict(case=row['case'], corpus=row['corpus'], sink=row.get('sink'),
        output_sha256=row.get('output_sha256'), source=normalized(row.get('source')))


def validate_row(row, samples, allocator):
    context = row['case'] + '/' + row['corpus']['shape']
    elapsed = OLD._validate_elapsed(row, samples, context)
    if row['case'] != 'cfb_open':
        OLD.verify_xls_row(row, samples, context, allocator)
    else:
        assert row.get('source') is None and row.get('sink') is None
        assert row.get('output_sha256') is None
        OLD.verify_operation_metrics(row, samples, context, allocator, elapsed['sample_order'])
    allocation = row['operation_metrics']['allocation']
    if allocator:
        vectors = {k: v['values'] for k, v in allocation.items() if isinstance(v, dict) and 'values' in v}
        for i in range(samples):
            assert vectors['failed_allocation_calls'][i] == 0
            assert vectors['live_bytes_after'][i] == vectors['live_bytes_before'][i] + vectors['allocated_bytes'][i] - vectors['deallocated_bytes'][i]
            assert vectors['region_peak_live_bytes'][i] >= max(vectors['live_bytes_before'][i], vectors['live_bytes_after'][i])
    else:
        assert allocation['status'] == 'unavailable'
    identity(row)
    return row


def validate_report(path, job, kind='normal'):
    report = read(path)
    meta = read(FOLDER / ('binary-' + kind + '.json'))
    OLD.verify_report_identity(report, {'binary_sha256': meta['sha256']},
        job['samples'], job['warmup'], kind == 'alloc', path.name)
    env = report['environment']
    assert env['git_revision'] == read(HERE / 'plan.json')['revision']
    assert env['cpu_affinity'] == str(read(HERE / 'plan.json')['cpu'])
    assert meta['path'] == str(SCRATCH / kind)
    assert report['binary_identity']['path'] == str(TARGET / 'retained-binaries' / kind)
    assert report['binary_identity']['binary_bytes'] == meta['bytes']
    selection = job['selection']
    assert report['configuration']['cases'] == selection['cases']
    if 'shapes' in selection:
        assert report['configuration']['corpus_shapes'] == selection['shapes']
        assert report['configuration']['payload_kinds'] == [selection['payload']]
    rows = report['results']
    expected = {(case, shape) for case in selection['cases'] for shape in selection.get('shapes', ['xls-source-backed'])}
    # XLS uses one fixed semantic corpus; its exact shape/hash is validated by
    # corpus binding and the repeated complete identity, not a guessed label.
    actual = [(r['case'], r['corpus']['shape'] if 'shapes' in selection else 'xls-source-backed') for r in rows]
    assert len(actual) == len(set(actual)) and set(actual) == expected
    catalog = read(path.with_name(path.name.replace('.json', '.catalog.json')))
    OLD.validate_binding(report, catalog)
    return [validate_row(row, job['samples'], kind == 'alloc') for row in rows]


def analyze():
    plan = read(HERE / 'plan.json')
    metadata = {kind: binary(kind) for kind in ('normal', 'alloc')}
    rows = []
    identities = {}
    rss = []
    for lane in ('native', 'alloc'):
        for job in jobs(lane):
            kind = 'alloc' if lane == 'alloc' else 'normal'
            r = receipt(job['name'], kind)
            command = r['command']
            assert command[:3] == ['taskset', '-c', str(plan['cpu'])]
            assert str(SCRATCH / kind) in command
            assert command[command.index('--case')+1] == ','.join(job['selection']['cases'])
            assert command[command.index('--samples')+1] == str(job['samples'])
            assert command[command.index('--warmup')+1] == str(job['warmup'])
            captured = validate_report(FOLDER / (job['name'] + '.json'), job, kind)
            if lane == 'native':
                value = read(FOLDER / (job['name'] + '.rss.json'))
                assert value['max_rss_kib'] > 0
                rss.append(dict(name=job['name'], **value))
            for row in captured:
                key = row['case'], row['corpus']['shape']
                ident = identity(row)
                if key in identities:
                    assert identities[key] == ident, key
                else:
                    identities[key] = ident
                item = dict(lane=lane, repeat=job['repeat'], case=key[0], shape=key[1],
                    report=job['name'] + '.json', samples=job['samples'], identity_equal=True,
                    elapsed_ns={m: row['elapsed_ns'][m] for m in METRICS})
                if lane == 'native':
                    item['operations_per_second'] = 1e9 / row['elapsed_ns']['mean']
                    item['standard_deviation_ns'] = row['elapsed_ns']['standard_deviation']
                    item['confidence_interval_95'] = row['elapsed_ns']['confidence_interval_95']
                else:
                    a = row['operation_metrics']['allocation']
                    item['allocation'] = {f: a[f]['values'] for f in OLD.ALLOCATOR_VECTOR_FIELDS}
                    item['allocation']['incremental_region_peak_live_bytes'] = [p-b for p,b in zip(a['region_peak_live_bytes']['values'],a['live_bytes_before']['values'])]
                    del item['elapsed_ns']
                rows.append(item)
    drift = []
    native = {(r['case'],r['shape'],r['repeat']): r for r in rows if r['lane']=='native'}
    for case,shape in sorted(identities):
        a,b = [native[case,shape,repeat] for repeat in (1,2)]
        for stat in METRICS:
            before,after = a['elapsed_ns'][stat],b['elapsed_ns'][stat]
            change = (after/before-1)*100
            if abs(change) > plan['review']['same_build_adverse_percent']:
                drift.append(dict(case=case,shape=shape,metric=stat,first=before,second=after,change_percent=change))
    for group in plan['groups']:
        a,b=[next(r for r in rss if r['name']==f'native-r{repeat}-{group}') for repeat in (1,2)]
        change=(b['max_rss_kib']/a['max_rss_kib']-1)*100
        if abs(change)>5:
            drift.append(dict(group=group,metric='whole_child_max_rss_kib',first=a['max_rss_kib'],second=b['max_rss_kib'],change_percent=change))
    return dict(status='pass',scope=plan['scope'],plan_sha256=sha(HERE/'plan.json'),
        native_samples=sum(r['samples'] for r in rows if r['lane']=='native'),
        allocation_samples=sum(r['samples'] for r in rows if r['lane']=='alloc'),
        rows=rows,identities=list(identities.values()),whole_child_rss=rss,
        same_build_variations_over_five_percent=drift,
        helpers={str(HELPER.relative_to(HERE.parent)):sha(HELPER),
            'tools/summarize_crud_baseline.py':sha(REPO/'tools/summarize_crud_baseline.py'),
            'tools/validate_perf_corpus_binding.py':sha(REPO/'tools/validate_perf_corpus_binding.py')},
        limits='No before/after production speedup, physical I/O/cold/range/native-producer/scaling claim. Shared-host two-child replication; instrumented allocation timings excluded.')


if __name__ == '__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('output',nargs='?',type=Path)
    args=parser.parse_args()
    output=args.output or HERE/'analysis.json'
    output.write_text(json.dumps(analyze(),indent=2)+'\n')
    print('Numerical evidence verified:',output)
