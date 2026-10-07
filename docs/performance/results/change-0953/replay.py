#!/usr/bin/env python3
"""Relocated offline replay adapter for the sealed 0953 FP packet.

The frozen v3 reader already accepts --packet, but captured JSON descriptors
contain absolute paths from the original packet.  This adapter loads the exact
reader and semantic helper from a relocated extraction, rewrites only existing
packet-local path strings in JSON, and leaves the explicitly bound diagnostic
binary at its recorded external target path.  It never starts a workload or
perf.
"""
from __future__ import annotations
import argparse
import importlib.util
import json
import sys
from pathlib import Path
from typing import Any

HIST = Path('/home/zhuhe/litchi-performance-0953').resolve()
MOVED = Path('/home/zhuhe/litchi-performance-0953-replay-v1').resolve()
DEFAULT_READER = MOVED / 'tools/analyze_0953_v3.py'

def map_string(value: str) -> str:
    prefix = str(HIST)
    if value == prefix or value.startswith(prefix + '/'):
        suffix = value[len(prefix):].lstrip('/')
        candidate = MOVED / suffix
        # The target binary is intentionally an external immutable dependency;
        # only packet-local paths whose relocated counterpart exists are moved.
        if candidate.exists():
            return str(candidate)
    return value

def relocate_json(value: Any) -> Any:
    if isinstance(value, dict):
        return {key: relocate_json(item) for key, item in value.items()}
    if isinstance(value, list):
        return [relocate_json(item) for item in value]
    if isinstance(value, str):
        return map_string(value)
    return value

def load_module(path: Path, name: str) -> Any:
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f'cannot load {path}')
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module

def configure(reader: Any) -> Any:
    old_load = reader.load_json
    def relocated_load(path: Path) -> Any:
        return relocate_json(old_load(path))
    reader.load_json = relocated_load
    old_import_semantic = reader.import_semantic
    def relocated_import(path: Path) -> Any:
        sem = old_import_semantic(path)
        old_sem_load = sem.read_json
        def sem_load(sem_path: Path) -> Any:
            return relocate_json(old_sem_load(sem_path))
        sem.read_json = sem_load
        return sem
    reader.import_semantic = relocated_import
    return reader

def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--packet', type=Path, default=MOVED)
    parser.add_argument('--plan', type=Path, default=MOVED / 'plan-0953-v3.json')
    parser.add_argument('--reader', type=Path, default=DEFAULT_READER)
    parser.add_argument('--index', type=int, required=True)
    parser.add_argument('--sample-export', type=Path, required=True)
    parser.add_argument('--mmap-export', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    reader = configure(load_module(args.reader.resolve(), 'analyze_0953_relocated'))
    output = args.output.resolve()
    try:
        result, code = reader.run(argparse.Namespace(
            packet=args.packet.resolve(), plan=args.plan.resolve(), index=args.index,
            sample_export=args.sample_export.resolve(), mmap_export=args.mmap_export.resolve(),
            output=output))
        reader.write_exclusive(output, result)
        print(json.dumps({'status':result['status'],'code':code,'output':str(output),
                          'sha256':reader.sha(output)}, sort_keys=True))
        return code
    except (reader.AdmissionError, AssertionError, OSError, ValueError, KeyError,
            ImportError, reader.struct.error) as exc:
        failure={'schema':'litchi.performance.0953.native-admission.v3','status':'error',
                 'error':str(exc),'scope':'Relocated replay failed closed; no admission.'}
        try:
            reader.write_exclusive(output, failure)
        except Exception:
            pass
        print(f'replay_0953_v1: {exc}', file=sys.stderr)
        return 2

if __name__ == '__main__': raise SystemExit(main())
