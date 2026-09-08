#!/usr/bin/env python3
"""Portable cross-checks for final profiling, validation and runtime custody."""
import gzip
import hashlib
import importlib.util


def check(root, *, complete):
    import analyze
    import verify as v

    def read(name):
        return v.read_json(v.bundle_file(root, name, name), name)

    def check_meta(name, expected):
        v.check_metadata(v.bundle_file(root, name, name), expected, name)

    binaries = read('binaries.json')
    profiles = read('profiles.json')
    v.require(profiles['schema'] == 'docx-plain-paragraph-tail-append-profiles-v1', 'profile schema')
    v.require(profiles['status'] == 'pass' and profiles['formal_samples'] is False, 'profile status/scope')
    v.require(profiles['binary'] == binaries['normal'], 'profile binary differs')
    records = profiles['records']
    v.require(set(records) == {'perf-stat', 'perf-record', 'strace', 'perf-report', 'perf-script', 'raw_profile'}, 'profile inventory')
    for name in ('perf-stat', 'perf-record', 'strace', 'perf-report', 'perf-script'):
        item = records[name]
        path = f'validation/profile-{name}.json'
        v.require(item['path'] == path and item['exit_code'] == 0, f'profile {name} failed')
        v.require(v.sha256_file(root / path)[0] == item['sha256'], f'profile {name} receipt digest')
        receipt = read(path)
        v.require(receipt['exit_code'] == 0 and receipt['source_unchanged'] is True, f'profile {name} receipt status')
        v.require(receipt['source_after']['sha256'] == binaries['normal']['source_manifest_sha256'], f'profile {name} source')
        if name in ('perf-stat', 'perf-record', 'strace'):
            argv = receipt['argv']
            expected_tail = [binaries['normal']['path'], '--mode', 'total', '--counts', '131072', '--samples', '30', '--warmups', '3', '--json']
            v.require(argv[-11:-1] == expected_tail, f'profile {name} operation command')
            v.require(argv[-1].endswith(f'/profiles/{name}.report.json'), f'profile {name} output command')
            checked = analyze.validate_report(read(f'profiles/{name}.report.json'),
                                              {'count': 131072, 'mode': 'total', 'instrumentation': 'normal'}, name)
            v.require(checked['corpus'] == read('corpus-manifest.json')['cases']['131072'], f'profile {name} corpus')
    packed = v.bundle_file(root, 'profiles/perf.data.gz', 'packed perf data')
    v.check_metadata(packed, records['raw_profile']['compressed'], 'packed perf data')
    digest = hashlib.sha256()
    size = 0
    with gzip.open(packed, 'rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            digest.update(block)
            size += len(block)
    v.require({'bytes': size, 'sha256': digest.hexdigest()} == records['raw_profile']['uncompressed'], 'perf data decompression identity')
    spec = importlib.util.spec_from_file_location('docx_append_profile_analysis', root / 'profile-analysis.py')
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    v._compare(module.derive(root), read('profile-summary.json'), 'profile summary')
    if not complete:
        return {'profiles': 'pass'}

    live = read('live-state.json')
    cleanup = read('cleanup.json')
    initial = read('initial-state.json')
    v.require(live['status'] == cleanup['status'] == 'pass', 'live/cleanup failed')
    v.require(live['binaries'] == cleanup['removed_binaries'] == binaries, 'live/cleanup binary identities')
    v.require(live['source']['sha256'] == binaries['normal']['source_manifest_sha256'], 'live final source')
    v._manifest(root, live['source'], 'live source')
    v.require(live['embedded_input_count'] == 8, 'live embedded input count')
    v.require(live['protected_files'] == cleanup['protected_files'] == initial['protected_files'], 'protected file custody')
    v.require(cleanup['live_state_sha256'] == v.sha256_file(root / 'live-state.json')[0], 'cleanup live binding')
    v.require(cleanup['temporary_root_absent'] is True, 'cleanup root absence record')
    v.require(len(cleanup['shared_caches_retained']) == 2, 'shared cache retention record')
    v.require(v.timestamp(cleanup['checked_utc'], 'cleanup time') >= v.timestamp(live['checked_utc'], 'live time'), 'cleanup chronology')

    evidence = read('evidence-validation.json')
    helpers = {p.name for p in root.glob('*.py')}
    v.require(set(evidence['helpers']) == helpers, 'evidence helper inventory')
    for name, expected in evidence['helpers'].items():
        check_meta(name, expected)
    check_meta('validation/evidence-tests-final.json', evidence['test_receipt'])
    v.require(read('validation/evidence-tests-final.json')['exit_code'] == 0, 'evidence tests failed')
    return {'profiles': 'pass', 'live_cleanup': 'pass', 'evidence_helpers': len(helpers)}
