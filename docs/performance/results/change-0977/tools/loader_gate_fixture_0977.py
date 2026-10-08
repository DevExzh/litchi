#!/usr/bin/env python3
"""Bounded offline exercise of the 0977 twelve-packet loader gate.

The packet files are synthetic JSON wrappers around one already-decoded v7
packet.  All large artifact bindings still point to the retained v7 child; no
capture, build, or decoder command is run here.  The first manifest entry is
also mutated once to verify fail-closed packet identity validation.
"""
from __future__ import annotations
import copy
import datetime as dt
import hashlib
import importlib.util
import json
import os
import sys
from pathlib import Path
from typing import Any

ROOT = Path('/home/zhuhe/litchi-performance-0977')
BASE = ROOT / 'integration-v7'
CHILD = BASE / 'capture' / 'child-01'
OUT = ROOT / 'loader-gate-v1'
PLAN_PATH = BASE / 'plan-integration.json'
BASE_PACKET = CHILD / 'child-01.packet.json'
RECEIPT = OUT / 'loader-gate-receipt-v1.json'


def sha(path: Path) -> str:
    h = hashlib.sha256()
    with path.open('rb') as f:
        for b in iter(lambda: f.read(1 << 20), b''):
            h.update(b)
    return h.hexdigest()


def binding(path: Path) -> dict[str, Any]:
    return {'path': str(path.absolute()), 'bytes': path.stat().st_size, 'sha256': sha(path)}


def load_module(name: str, path: Path) -> Any:
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f'cannot import {path}')
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


def write_json(path: Path, value: Any) -> None:
    if path.exists():
        raise RuntimeError(f'refusing to overwrite {path}')
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, indent=2, sort_keys=True, allow_nan=False) + '\n', encoding='utf-8')


def packet_for_index(base: dict[str, Any], index: int) -> dict[str, Any]:
    packet = copy.deepcopy(base)
    packet['packet_id'] = f'child-{index:02d}'
    packet['child_number'] = index
    # Keep only the bindings needed by _decode_packet plus the custody/loss
    # aliases.  This avoids rehashing optional decoder diagnostics while still
    # validating raw-task, corrected script/log, layout, runner command,
    # affinity observation, terminal, and report custody for every child.
    keep = {
        'child-01.affinity-observations.json', 'command', 'corrected_log',
        'corrected_script', 'launch', 'layout', 'loss.json', 'raw_task',
        'report', 'terminal', 'terminal-summary.json', 'terminal_summary',
        'terminal_report',
    }
    packet['artifacts'] = {k: v for k, v in packet['artifacts'].items() if k in keep}
    packet.pop('runner_bindings', None)
    return packet


def manifest(plan_path: Path, packet_paths: list[Path]) -> dict[str, Any]:
    return {
        'schema': 'litchi.performance.0977.capture_manifest.v1',
        'status': 'complete', 'no_retry': True, 'serial': True,
        'children_requested': 12, 'children_completed': 12,
        'plan': binding(plan_path),
        'packets': [str(p.absolute()) for p in packet_paths],
        'failure': None,
    }


def main() -> int:
    if OUT.exists():
        raise RuntimeError(f'output already exists: {OUT}')
    OUT.mkdir(parents=True)
    reducer = load_module('reduce_0977_loader_fixture', ROOT / 'tools' / 'reduce_0977.py')
    plan = json.loads(PLAN_PATH.read_text(encoding='utf-8'))
    base = json.loads(BASE_PACKET.read_text(encoding='utf-8'))
    paths = []
    for i in range(1, 13):
        path = CHILD / f'child-01.loader-{i:02d}.packet.json'
        write_json(path, packet_for_index(base, i))
        paths.append(path)
    manifest_path = OUT / 'capture-manifest.json'
    write_json(manifest_path, manifest(PLAN_PATH, paths))
    packets = reducer._load_packets(PLAN_PATH, manifest_path, plan)
    result = reducer._reduce_packets(PLAN_PATH, manifest_path, plan, packets)

    bad = packet_for_index(base, 99)
    bad_path = CHILD / 'child-01.loader-bad.packet.json'
    write_json(bad_path, bad)
    bad_manifest_path = OUT / 'capture-manifest-bad.json'
    bad_paths = [bad_path] + paths[1:]
    write_json(bad_manifest_path, manifest(PLAN_PATH, bad_paths))
    try:
        reducer._load_packets(PLAN_PATH, bad_manifest_path, plan)
    except Exception as exc:
        negative = {'status': 'rejected', 'exception': type(exc).__name__, 'message': str(exc)[:500]}
    else:
        raise RuntimeError('mutated first packet identity was accepted')

    receipt = {
        'schema': 'litchi.performance.0977.loader_gate_fixture.v1',
        'status': 'pass', 'created_utc': dt.datetime.now(dt.timezone.utc).isoformat(),
        'workload_started': False, 'build_started': False, 'perf_commands_started': False,
        'base_packet': binding(BASE_PACKET),
        'plan': binding(PLAN_PATH), 'manifest': binding(manifest_path),
        'loaded_children': len(packets),
        'reduced_children': len(result.get('children', [])),
        'positive_scope': 'twelve synthetic packet wrappers over one immutable decoded child; full reducer loader and aggregation gates exercised',
        'negative_mutation': negative,
        'limitation': 'This is manifest/loader custody and arithmetic validation; duplicated packets are not twelve native observations.',
    }
    write_json(RECEIPT, receipt)
    print(json.dumps({'status': 'pass', 'receipt': str(RECEIPT), 'sha256': sha(RECEIPT), 'loaded_children': len(packets), 'negative': negative}, indent=2, sort_keys=True))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
