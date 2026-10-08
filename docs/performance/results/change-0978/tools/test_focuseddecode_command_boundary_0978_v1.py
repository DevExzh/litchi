#!/usr/bin/env python3
"""Check the 0978 offline perf-reader custody boundary without running perf."""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path

ROOT = Path('/home/zhuhe/litchi-performance-0978')
PLAN = ROOT / 'integration-command-v1' / 'plan.json'
TOOLS = ROOT / 'tools'


def load_decoder():
    path = TOOLS / 'focuseddecode_0978.py'
    spec = importlib.util.spec_from_file_location('focuseddecode_command_boundary_0978', path)
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def main() -> int:
    decoder = load_decoder()
    plan = json.loads(PLAN.read_text())
    cfg = decoder._validate_plan(plan)
    expected_env = set(decoder._base_env())
    rows = []
    for child in ('child-01', 'child-02'):
        directory = ROOT / 'integration-command-v1' / child
        receipts = {
            'perf_header': directory / f'{child}.perf-header.command.json',
            'terminal_report': directory / f'{child}.terminal-report.command.json',
            'perf_evlist': directory / f'{child}.perf-evlist-v.command.json',
            'perf_mmap_dump': directory / f'{child}.perf-mmap-dump.command.json',
            'corrected': directory / f'{child}.corrected.command.json',
        }
        values = [json.loads(path.read_text()) for path in receipts.values()]
        assert all(value['schema'] == 'litchi.performance.0978.offline_command.v1' for value in values)
        assert all(value['executed'] is False and value['reused_existing'] is True for value in values)
        assert all(value['native_workload_started'] is False and value['returncode'] == 0 for value in values)
        assert set(values[0]['environment']) == expected_env
        assert set(values[-1]['environment']) >= expected_env | {'LD_PRELOAD', 'LITCHI_0973_SHIM_MODE', 'LITCHI_0973_SHIM_LOG'}
        rows.append({'child': child, 'receipt_kinds': [value['kind'] for value in values], 'all_reused': True})
    # A strict/fresh plan must reject an existing reader output even when a
    # valid retained-input receipt is present.  This prevents partial replay
    # artifacts from masquerading as a new child.
    value = json.loads((ROOT / 'integration-command-v1/child-01/child-01.perf-header.command.json').read_text())
    bad_cfg = dict(cfg)
    bad_cfg['analysis_only_replay'] = False
    try:
        decoder._reuse_or_fail(
            bad_cfg,
            ROOT / 'integration-command-v1/child-01/child-01.perf-header.command.json',
            kind='perf_header',
            argv=value['argv'],
            env=decoder._base_env(),
            output=ROOT / 'integration-command-v1/child-01/child-01.perf-header.txt',
            stderr=ROOT / 'integration-command-v1/child-01/child-01.perf-header.stderr',
        )
    except decoder.DecodeFailure as exc:
        assert exc.code == 'offline_command_receipt'
        negative = {'status': 'pass', 'code': exc.code}
    else:
        raise AssertionError('strict plan accepted retained offline output')
    print(json.dumps({'schema': 'litchi.performance.0978.command_boundary_controls.v1', 'status': 'pass', 'offline_only': True, 'focused_decoder': decoder.binding(TOOLS / 'focuseddecode_0978.py'), 'children': rows, 'strict_reuse_negative': negative}, sort_keys=True))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
