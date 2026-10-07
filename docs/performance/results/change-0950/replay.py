#!/usr/bin/env python3
"""Relocated 0950 replay adapter correcting the workload-cwd subdirectory binding.

The frozen v4 adapter is retained in the packet and its failed negative-control
attempt is preserved separately.  This derived adapter imports those frozen
bytes, maps the historical workload-cwd-v1 path to the extracted physical
directory, and leaves all reducer/analysis logic untouched.
"""
from __future__ import annotations
import argparse, hashlib, importlib.util, json
from pathlib import Path
from typing import Any

V4_REL = 'tools/replay_0950_v4.py'
V5_SCHEMA = 'litchi.performance.0950.portable-replay-adapter.v5'

def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()

def load_v4(packet: Path) -> Any:
    path = packet / V4_REL
    spec = importlib.util.spec_from_file_location('litchi_0950_frozen_replay_v4', path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f'cannot import {path}')
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    module.PACKET = packet.resolve()
    module.ADAPTER_PATH = path.resolve()
    return module

def install_corrected(module: Any, reducer: Any, historical: str):
    packet = reducer.PACKET.resolve()
    historical_root = Path(historical)
    original_command_for = reducer.command_for
    original_path_value = reducer.path_value

    def path_value(value: Any, label: str, base: Path = packet, *,
                   allow_absolute: bool = False) -> Any:
        physical = original_path_value(value, label, base=base,
                                        allow_absolute=allow_absolute)
        if label == 'workload_cwd':
            raw = Path(physical)
            if raw != historical_root and historical_root not in raw.parents:
                raise module.ReplayError(f'workload cwd escaped historical root: {raw}')
            relocated = packet / raw.relative_to(historical_root)
            return module.HistoricalCwd(relocated, str(raw))
        return physical

    def command_for(plan: dict[str, Any], job: dict[str, Any], folder: Path) -> list[str]:
        command = original_command_for(plan, job, folder)
        return [module.rewrite_packet_prefix(value, packet, historical)
                for value in command]

    reducer.path_value = path_value
    reducer.command_for = command_for

    def restore() -> None:
        reducer.path_value = original_path_value
        reducer.command_for = original_command_for
    return restore

def corrected_negative(module: Any, reducer: Any, historical: str, receipt: Path) -> dict[str, Any]:
    plan = reducer.load_plan()
    job = reducer.jobs('qualification')[0]
    folder = module.PACKET / 'qualification' / reducer.child_name(job)
    restore = install_corrected(module, reducer, historical)
    try:
        expected = reducer.command_for(plan, job, folder)
        workload = reducer.path_value(plan['workload_cwd'], 'workload_cwd', allow_absolute=True)
        physical_workload = str(workload.physical)
        captured_workload = str(workload)
    finally:
        restore()
    module.require(Path(physical_workload).resolve().is_relative_to(module.PACKET.resolve()),
                   'workload cwd was not relocated under extracted packet')
    module.require(Path(physical_workload).is_dir() and not Path(physical_workload).is_symlink(),
                   'relocated workload cwd is not a real directory')
    sample = reducer.read_json(folder / 'start.json')
    module.require(sample.get('cwd') == captured_workload,
                   'captured workload cwd does not match historical workload path')
    module.require(sample.get('argv') == expected,
                   'captured argv does not match corrected relocated expectation')
    tampered_cwd = dict(sample); tampered_cwd['cwd'] = captured_workload + '/tampered'
    try:
        module.invocation_check(tampered_cwd, expected, captured_workload)
    except module.ReplayError:
        cwd_result = 'rejected'
    else:
        raise module.ReplayError('negative control accepted tampered workload cwd')
    tampered_argv = dict(sample); tampered_argv['argv'] = list(expected)
    tampered_argv['argv'][-1] += '.tampered'
    try:
        module.invocation_check(tampered_argv, expected, captured_workload)
    except module.ReplayError:
        argv_result = 'rejected'
    else:
        raise module.ReplayError('negative control accepted tampered argv')
    missing = module.PACKET / '.replay-negative-missing-artifact-v5'
    module.require(not missing.exists() and not missing.is_symlink(), 'negative path unexpectedly exists')
    try:
        reducer.digest(missing)
    except reducer.ReviewError:
        missing_result = 'rejected'
    else:
        raise module.ReplayError('negative control accepted missing artifact')
    result = {
        'schema': V5_SCHEMA + '.negative-controls', 'status': 'pass',
        'derived_adapter': {'path': str(Path(__file__).resolve()), 'sha256': sha(Path(__file__).resolve())},
        'frozen_adapter': {'path': V4_REL, 'sha256': sha(module.ADAPTER_PATH)},
        'frozen_reducer': {'path': module.REDUCER_RELATIVE, 'sha256': module.REDUCER_SHA256},
        'historical_root': historical, 'historical_workload_cwd': captured_workload,
        'relocated_packet_root': str(module.PACKET.resolve()),
        'physical_workload_cwd': physical_workload,
        'controls': {'tampered_workload_cwd': cwd_result,
                     'tampered_argv': argv_result,
                     'missing_relocated_artifact': missing_result},
        'v4_failure': 'preserved separately: v4 compared workload cwd to packet root',
    }
    module.write_once(receipt, result)
    return result

def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument('action', choices=['negative-controls', 'all'])
    ap.add_argument('--packet', type=Path, required=True)
    ap.add_argument('--output', type=Path)
    ap.add_argument('--receipt', type=Path)
    a = ap.parse_args(); packet = a.packet.resolve(); module = load_v4(packet)
    historical, authorization = module.bind_historical_root()
    reducer = module.load_frozen_reducer()
    if a.action == 'negative-controls':
        receipt = a.receipt or packet / 'receipts/replay-adapter-negative-controls-v5.json'
        result = corrected_negative(module, reducer, historical, receipt)
        print(json.dumps({'status': result['status'], 'receipt': str(receipt)}, sort_keys=True)); return 0
    if a.output is None: raise RuntimeError('--output is required for all')
    # The frozen v4 replay() looks up its adapter by module-global name. Bind
    # that hook to the corrected workload-cwd mapping for the entire reducer
    # call, then restore the frozen function before returning.
    frozen_install = module.install_relocation_adapter
    module.install_relocation_adapter = lambda current_reducer, root: install_corrected(
        module, current_reducer, root)
    try:
        result = module.replay(reducer, historical, 'all', a.output.absolute())
    finally:
        module.install_relocation_adapter = frozen_install
    print(json.dumps({'status': result['status'], 'stage': result['stage'],
                      'historical_root': historical, 'authorization': authorization}, sort_keys=True)); return 0

if __name__ == '__main__':
    raise SystemExit(main())
