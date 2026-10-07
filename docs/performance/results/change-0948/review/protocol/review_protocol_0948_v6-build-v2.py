#!/usr/bin/env python3
"""Focused offline preflight for the 0948 v2 build-receipt bridge.

Constructs the same in-memory plan that freeze_0948_v4.py would write, then
runs the v2 terminal/reuse validators against the actual v3 build receipts and
retained binary copies.  It never writes plan-v2.json, launches a build, or
starts a workload.  Prior gate controls are linked, not rerun.
"""
from __future__ import annotations

import hashlib
import importlib.util
import json
import os
from pathlib import Path
from typing import Any

PACKET = Path('/home/zhuhe/litchi-performance-0948')
TOOLS = PACKET / 'tools'
FREEZE = TOOLS / 'freeze_0948_v4.py'
REUSE = TOOLS / 'reuse_0948_v2.py'
CAPTURE = TOOLS / 'capture_0948_v2.py'
REDUCE = TOOLS / 'reduce_0948_v2.py'
TEMPLATE = PACKET / 'plan-v1-draft.json'
PRIOR_GATE = PACKET / 'review/protocol/review_protocol_0948_v4-final.json'
OUT = PACKET / 'review/protocol/review_protocol_0948_v6-build-v2.json'
EXPECTED = {
    'freeze': '84db5881fbfc4aeda4b3f86b850a6f29b0ace3832f9abbc6dfb12f141bd4041d',
    'reuse': 'c83c1270e42a82d73bc9f08f77f7cb2ac9087ff969b1e3082f8b6f84699c726a',
    'capture': '307d49f623368860f87809ec28862448cc69be5565dcd2cb0f3a2da4eefcdeb1',
    'reduce': '7f5f3abdbbddd78138359fa824f25611ecc113d4e89bb3477b6cd874d025a630',
    'prior_gate': 'ecbbf311342afe1affe6d577fffb23caa7be74ddf9b00690306c09e2adf8a7ca',
}


def sha(path: Path) -> str:
    h = hashlib.sha256()
    with path.open('rb') as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b''):
            h.update(chunk)
    return h.hexdigest()


def load(path: Path, name: str) -> Any:
    spec = importlib.util.spec_from_file_location(name, path)
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def strict(path: Path) -> Any:
    return json.loads(path.read_text(encoding='utf-8'))


def main() -> None:
    checks: list[dict[str, Any]] = []
    helper_bindings = {}
    for name, path in {'freeze': FREEZE, 'reuse': REUSE, 'capture': CAPTURE,
                       'reduce': REDUCE, 'prior_gate': PRIOR_GATE}.items():
        actual = sha(path)
        helper_bindings[name] = {'path': str(path), 'bytes': path.stat().st_size,
                                 'sha256': actual, 'expected_sha256': EXPECTED[name]}
        checks.append({'id': f'hash_{name}', 'status': 'pass' if actual == EXPECTED[name] else 'fail'})

    freeze = load(FREEZE, 'freeze_0948_v4_review')
    reuse = load(REUSE, 'reuse_0948_v2_review')
    capture = load(CAPTURE, 'capture_0948_v2_review')
    reduce = load(REDUCE, 'reduce_0948_v2_review')
    template = strict(TEMPLATE)
    plan = dict(template)
    plan['source'] = freeze.source_binding()
    plan['build_revisions'] = dict(freeze.BUILD_REVISIONS)
    plan['build_binding_ids'] = dict(freeze.BUILD_BINDING_IDS)
    plan['roles'] = freeze.role_bindings()
    plan['semantic_validator'] = freeze.descriptor(TOOLS / 'semantic_0948_v1.py')
    plan['helpers'] = {
        'capture': freeze.descriptor(CAPTURE),
        'reducer': freeze.descriptor(REDUCE),
        'semantic': freeze.descriptor(TOOLS / 'semantic_0948_v1.py'),
        'freeze': freeze.descriptor(FREEZE),
        'method': freeze.descriptor(PACKET / 'METHOD-v2-draft.md'),
        'reuse_validator': freeze.descriptor(REUSE),
    }

    checks.append({'id': 'in_memory_freeze_source_binding', 'status': 'pass',
                   'candidate_files': plan['source']['candidate_files'],
                   'file_count': plan['source']['file_count']})
    checks.append({'id': 'role_bindings_use_retained_binary_copies',
                   'status': 'pass' if all(
                       value['path'].startswith('binaries/') and
                       not value['path'].startswith('target-') and
                       value['build_receipt'].startswith(('receipts/', 'builds/'))
                       for role in plan['roles'].values() for value in role.values()) else 'fail',
                   'paths': {role: {kind: value['path'] for kind, value in kinds.items()}
                             for role, kinds in plan['roles'].items()}})

    validated = {}
    for kind in ('native', 'allocator'):
        raw_path = PACKET / f'builds/candidate-{kind}-v3/terminal.json'
        raw = strict(raw_path)
        binary = plan['roles']['candidate'][kind]
        try:
            reuse.validate_build_terminal_0948(raw, 'candidate', kind, plan, binary=binary)
        except Exception as error:
            checks.append({'id': f'candidate_{kind}_terminal_v2_validator', 'status': 'fail',
                           'error': type(error).__name__ + ': ' + str(error)})
        else:
            checks.append({'id': f'candidate_{kind}_terminal_v2_validator', 'status': 'pass',
                           'terminal_sha256': sha(raw_path), 'returncode': raw['returncode'],
                           'binary_sha256': binary['sha256']})
            validated[kind] = {'terminal': str(raw_path), 'terminal_sha256': sha(raw_path),
                               'binary': binary}

    reuse_raw = strict(PACKET / 'receipts/baseline-reuse-v2.json')
    baseline_results = {}
    for kind in ('native', 'allocator'):
        try:
            reuse.validate_baseline_reuse_0948(reuse_raw, plan, kind,
                                               binary=plan['roles']['baseline'][kind])
        except Exception as error:
            checks.append({'id': f'baseline_{kind}_reuse_v2_validator', 'status': 'fail',
                           'error': type(error).__name__ + ': ' + str(error)})
        else:
            checks.append({'id': f'baseline_{kind}_reuse_v2_validator', 'status': 'pass',
                           'binary_sha256': plan['roles']['baseline'][kind]['sha256']})
            baseline_results[kind] = 'pass'

    # The v2 capture/reducer loaders must resolve the v2 reuse helper.  This is
    # an import/callsite check only; neither module is asked to analyze a packet.
    checks.append({'id': 'capture_imports_reuse_v2',
                   'status': 'pass' if capture._reuse_validator().__file__ == str(REUSE) else 'fail',
                   'resolved': capture._reuse_validator().__file__})
    checks.append({'id': 'reduce_imports_reuse_v2',
                   'status': 'pass' if reduce._reuse_validator().__file__ == str(REUSE) else 'fail',
                   'resolved': reduce._reuse_validator().__file__})

    all_pass = all(item['status'] == 'pass' for item in checks)
    receipt = {
        'schema': 'litchi.performance.0948.build-v2-terminal-preflight.v1',
        'status': 'pass' if all_pass else 'fail',
        'scope': 'actual v3 candidate terminals and retained baseline reuse; no plan write, build, or workload',
        'plan': {'source': 'in-memory freeze_0948_v4 construction', 'status': 'not written',
                 'future_final_binding_required': True},
        'helper_bindings': helper_bindings,
        'validated_candidate': validated,
        'validated_baseline': baseline_results,
        'prior_gate_controls': {'path': str(PRIOR_GATE), 'sha256': sha(PRIOR_GATE), 'status': 'reused; not rerun'},
        'checks': checks,
    }
    data = (json.dumps(receipt, indent=2, sort_keys=True, allow_nan=False) + '\n').encode()
    if OUT.exists():
        raise RuntimeError(f'refusing to overwrite {OUT}')
    with OUT.open('xb') as stream:
        stream.write(data)
        stream.flush()
        os.fsync(stream.fileno())
    print(json.dumps({'status': receipt['status'], 'output': str(OUT),
                      'bytes': len(data), 'sha256': sha(OUT), 'checks': checks}, sort_keys=True))


if __name__ == '__main__':
    main()
