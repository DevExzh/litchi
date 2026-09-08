#!/usr/bin/env python3
"""Check final runtime-cleanup and helper custody without opening runtime paths."""
import hashlib
import json


def check(root, complete):
    def require(condition, message):
        if not condition:
            raise ValueError(message)

    def read(name):
        path = root / name
        require(path.is_file() and not path.is_symlink(), f'missing regular {name}')
        return json.loads(path.read_text())

    def meta(name):
        path = root / name
        return {'bytes': path.stat().st_size, 'sha256': hashlib.sha256(path.read_bytes()).hexdigest()}

    if not complete:
        return {'complete': False}
    binaries = {f'{arm}-{mode}': binary for arm in ('control', 'candidate')
                for mode, binary in read(f'{arm}-binaries.json').items()}
    live = read('live-state.json')
    cleanup = read('cleanup.json')
    initial = read('initial-state.json')
    require(live['status'] == cleanup['status'] == 'pass', 'live/cleanup failed')
    require(live['binaries'] == cleanup['removed_binaries'] == binaries, 'live/cleanup binary identity')
    require(live['source']['sha256'] == binaries['candidate-normal']['source_manifest_sha256'], 'live candidate source')
    require(live['protected_files'] == cleanup['protected_files'] == initial['protected_files'], 'protected file custody')
    require(cleanup['live_state_sha256'] == meta('live-state.json')['sha256'], 'cleanup live binding')
    require(live['embedded_input_count'] == 8 and cleanup['temporary_root_absent'] is True, 'cleanup input count/root absence')
    require(len(cleanup['shared_caches_retained']) == 2, 'shared cache retention')
    evidence = read('evidence-validation.json')
    helpers = {path.name for path in root.glob('*.py')}
    require(set(evidence['helpers']) == helpers, 'helper inventory')
    for name, expected in evidence['helpers'].items():
        require(meta(name) == expected, f'helper digest differs: {name}')
    require(meta('validation/evidence-tests-final.json') == evidence['test_receipt'], 'evidence test binding')
    require(read('validation/evidence-tests-final.json')['exit_code'] == 0, 'evidence tests failed')
    return {'complete': True, 'live_cleanup': 'pass', 'helpers': len(helpers)}
