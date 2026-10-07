#!/usr/bin/env python3
"""Validate a completed 0949 offline replay without rerunning any tool."""
from __future__ import annotations

import argparse
import hashlib
import json
import os
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

LABELS = (
    ('baseline-medium', 'h00-baseline-primary-percent-medium'),
    ('candidate-medium', 'h01-candidate-primary-percent-medium'),
    ('baseline-dense', 'h02-baseline-primary-percent-dense-sparse'),
    ('candidate-dense', 'h03-candidate-primary-percent-dense-sparse'),
)
AGGREGATES = ('size-aggregates.tsv', 'trace-aggregates.jsonl', 'trace-size-aggregates.tsv')
ATTR_OUTPUTS = ('window-traces.jsonl', 'window-physical-size.histogram', 'window-inline-size.histogram',
                'window-physical-stacks.folded', 'window-inline-stacks.folded', 'publication-symbol-scoped.json')


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open('rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            h.update(block)
    return h.hexdigest()


def descriptor(path: Path) -> dict[str, Any]:
    if not path.is_file() or path.is_symlink():
        raise RuntimeError(f'not a regular file: {path}')
    return {'path': str(path.resolve()), 'bytes': path.stat().st_size, 'sha256': sha256(path)}


def load(path: Path) -> Any:
    with path.open(encoding='utf-8') as stream:
        return json.load(stream)


def normalized(value: Any, key: str = '') -> Any:
    if isinstance(value, dict):
        if key in {'decoder_summary', 'summary'} and {'bytes', 'sha256'} <= set(value):
            return {'relocated_summary_descriptor': True}
        if key == 'outputs':
            return {name: '<relocated-output>' for name in sorted(value)}
        return {k: normalized(v, k) for k, v in sorted(value.items())
                if k not in {'path', 'root', 'output_root', 'started_utc', 'ended_utc', 'created_utc', 'generated_utc'}}
    if isinstance(value, list):
        return [normalized(item, key) for item in value]
    if isinstance(value, str) and (key.endswith('_path') or key in {'cwd', 'command'}):
        return '<relocated>'
    return value


def semantic_equal(left: Path, right: Path) -> dict[str, Any]:
    a = normalized(load(left)); b = normalized(load(right))
    return {'original': descriptor(left), 'replay': descriptor(right), 'normalized_equal': a == b}


def exact_equal(left: Path, right: Path, kind: str) -> dict[str, Any]:
    a = descriptor(left); b = descriptor(right)
    return {'kind': kind, 'original': a, 'replay': b,
            'exact_equal': a['bytes'] == b['bytes'] and a['sha256'] == b['sha256']}


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument('--packet-root', type=Path, required=True)
    ap.add_argument('--output-root', type=Path, required=True)
    ap.add_argument('--receipt', type=Path, required=True)
    args = ap.parse_args()
    packet = args.packet_root.resolve(); output = args.output_root.resolve(); receipt = args.receipt.resolve()
    if not packet.is_dir() or not output.is_dir() or receipt.exists() or receipt.is_symlink():
        raise SystemExit('packet/output must exist and receipt must be new')
    decoder_root = output / 'raw-decoder-replay-v1'
    attr_root = output / 'attribution-replay-v1'
    staging = output / 'reconciliation-input' / 'raw-decoder-v1'
    decoder_terminal = decoder_root / 'decoder-terminal.json'
    attr_terminal = attr_root / 'attribution-terminal.json'
    if load(decoder_terminal).get('status') != 'terminal' or load(attr_terminal).get('status') != 'terminal':
        raise SystemExit('offline stages are not terminal-success records')
    semantic: list[dict[str, Any]] = []
    exact: list[dict[str, Any]] = []
    hardlinks: list[dict[str, Any]] = []
    for label, _child in LABELS:
        semantic.append(semantic_equal(packet / 'raw-decoder-v1' / label / 'summary.json', decoder_root / label / 'summary.json'))
        semantic.append(semantic_equal(packet / 'attribution-final-v1' / label / 'summary.json', attr_root / label / 'summary.json'))
        for name in AGGREGATES:
            exact.append(exact_equal(packet / 'raw-decoder-v1' / label / name, decoder_root / label / name, 'fresh_decoder_aggregate'))
            # Reconciliation must consume the fresh output, not retained packet rows.
            fresh = decoder_root / label / name
            staged = staging / label / name
            hardlinks.append({'label': label, 'name': name, 'samefile': os.path.samefile(fresh, staged)})
        for name in ATTR_OUTPUTS:
            exact.append(exact_equal(packet / 'attribution-final-v1' / label / name, attr_root / label / name, 'attribution_output'))
    semantic.append(semantic_equal(packet / 'review/xml-multiwindow-source-profile-reconciliation-v2.json',
                                   output / 'xml-multiwindow-reconciliation-v2.json'))
    dependency = packet.parent.parent / 'litchi-performance-0949-archive' / 'dependency-0949-v1.json'
    # The archive workspace is a sibling of the relocated packet in the normal
    # layout; retain the explicit dependency path even if validation is moved.
    dependency = Path('/home/zhuhe/litchi-performance-0949-archive/dependency-0949-v1.json')
    receipt_data = {
        'schema': 'litchi.performance.0949.offline-relocated-replay-validation.v2',
        'status': 'pass' if all(x['normalized_equal'] for x in semantic) and
                  all(x['exact_equal'] for x in exact) and all(x['samefile'] for x in hardlinks) else 'fail',
        'recorded_utc': datetime.now(timezone.utc).isoformat(),
        'packet_root': str(packet), 'output_root': str(output),
        'decoder_terminal': descriptor(decoder_terminal),
        'attribution_terminal': descriptor(attr_terminal),
        'semantic_comparisons': semantic,
        'exact_comparisons': exact,
        'fresh_decoder_hardlink_checks': hardlinks,
        'reconciliation_input': str(staging.parent),
        'reconciliation_used_fresh_decoder_output': all(x['samefile'] for x in hardlinks),
        'dependency_binding': descriptor(dependency),
        'workload_launched': False, 'binary_launched': False, 'profiler_launched': False,
        'rerun_workload': False,
    }
    receipt.parent.mkdir(parents=True, exist_ok=True)
    receipt.write_text(json.dumps(receipt_data, indent=2, sort_keys=True) + '\n')
    if receipt_data['status'] != 'pass':
        raise SystemExit('offline replay validation failed; receipt retained')
    print(json.dumps({'status':'pass','receipt':str(receipt),'semantic':len(semantic),'exact':len(exact),'hardlinks':len(hardlinks)},sort_keys=True))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
