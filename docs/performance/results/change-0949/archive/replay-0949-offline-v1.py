#!/usr/bin/env python3
"""Relocated, workload-free 0949 decoder/attributor/reconciler replay.

The packet supplies retained heaptrack profiles and existing decoder outputs.
This helper launches only the reviewed offline Python tools; it never invokes a
benchmark, a binary, heaptrack, heaptrack_print, Cargo, or a source build.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

LABELS = (
    ('baseline-medium', 'h00-baseline-primary-percent-medium'),
    ('candidate-medium', 'h01-candidate-primary-percent-medium'),
    ('baseline-dense', 'h02-baseline-primary-percent-dense-sparse'),
    ('candidate-dense', 'h03-candidate-primary-percent-dense-sparse'),
)


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


def run_offline(command: list[str], cwd: Path, log_root: Path, name: str) -> dict[str, Any]:
    result = subprocess.run(command, cwd=cwd, check=False, capture_output=True, text=True)
    stdout = log_root / f'{name}.stdout'
    stderr = log_root / f'{name}.stderr'
    stdout.write_text(result.stdout)
    stderr.write_text(result.stderr)
    return {
        'name': name,
        'argv': command,
        'cwd': str(cwd.resolve()),
        'exit_code': result.returncode,
        'stdout': descriptor(stdout),
        'stderr': descriptor(stderr),
        'workload_launched': False,
    }


def read_json(path: Path) -> Any:
    with path.open(encoding='utf-8') as stream:
        return json.load(stream)


def normalized(value: Any, key: str = '') -> Any:
    # Paths, timestamps, and output roots are expected to move; counters,
    # hashes, labels, and all attribution values must remain equal.
    if isinstance(value, dict):
        if key in {'decoder_summary', 'summary'} and {'bytes', 'sha256'} <= set(value):
            # Decoder summaries embed output paths in their own JSON and thus
            # legitimately change byte count/hash after relocation. The raw
            # TSV/JSONL aggregate files are compared exactly below.
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


def compare_json(original: Path, replay: Path) -> dict[str, Any]:
    left = normalized(read_json(original))
    right = normalized(read_json(replay))
    return {
        'original': descriptor(original),
        'replay': descriptor(replay),
        'normalized_equal': left == right,
    }


def compare_exact(original: Path, replay: Path) -> dict[str, Any]:
    left = descriptor(original)
    right = descriptor(replay)
    return {
        'kind': 'exact_file_sha256',
        'original': left,
        'replay': right,
        'exact_equal': left['bytes'] == right['bytes'] and left['sha256'] == right['sha256'],
    }


def hardlink_decoder_tree(decoder_root: Path, staging_root: Path, labels: tuple[tuple[str, str], ...]) -> None:
    """Expose fresh decoder files under the reconciler's packet-shaped input.

    Hard links avoid a second 0949 copy while making the reconciler consume the
    freshly decoded aggregate files rather than the retained packet outputs.
    """
    raw_root = staging_root / 'raw-decoder-v1'
    raw_root.mkdir(parents=True, exist_ok=False)
    for label, _child in labels:
        source_dir = decoder_root / label
        target_dir = raw_root / label
        target_dir.mkdir()
        for name in ('summary.json', 'size-aggregates.tsv', 'trace-aggregates.jsonl',
                     'trace-size-aggregates.tsv'):
            source = source_dir / name
            target = target_dir / name
            if not source.is_file() or source.is_symlink():
                raise RuntimeError(f'missing fresh decoder aggregate: {source}')
            os.link(source, target)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument('--packet-root', type=Path, required=True,
                        help='relocated extracted 0949 packet')
    parser.add_argument('--source-root', type=Path, required=True,
                        help='external 0948 candidate-build-v3 source snapshot')
    parser.add_argument('--output-root', type=Path, required=True,
                        help='new external output root')
    args = parser.parse_args()
    packet = args.packet_root.resolve()
    source = args.source_root.resolve()
    output = args.output_root.resolve()
    if not packet.is_dir() or packet.is_symlink():
        raise SystemExit(f'packet root is not a real directory: {packet}')
    if not source.is_dir() or source.is_symlink():
        raise SystemExit(f'source root is not a real directory: {source}')
    if output.exists() or output.is_symlink():
        raise SystemExit(f'refusing existing output root: {output}')
    output.mkdir(parents=True, exist_ok=False)
    logs = output / 'tool-terminals'
    logs.mkdir()
    started = datetime.now(timezone.utc).isoformat()
    profiles: list[str] = []
    histograms: list[str] = []
    for label, child in LABELS:
        profile = packet / 'capture-v1' / 'children' / child / 'heaptrack.data.gz.zst'
        histogram = packet / 'heaptrack-exports-v2' / child / 'allocation-size.histogram'
        if not profile.is_file() or profile.is_symlink():
            raise SystemExit(f'missing retained profile: {profile}')
        if not histogram.is_file() or histogram.is_symlink():
            raise SystemExit(f'missing retained histogram: {histogram}')
        profiles.append(f'{label}={profile}')
        histograms.append(f'{label}={histogram}')

    decoder = packet / 'tools' / 'decode_heaptrack_0949_v1.py'
    attributor = packet / 'tools' / 'attribute_heaptrack_0949_v1.py'
    config = packet / 'tools' / 'attribution-config-root-v1.json'
    reconciler = packet / 'review' / 'xml-multiwindow-source-profile-reconciliation-v2.py'
    for path in (decoder, attributor, config, reconciler):
        if not path.is_file() or path.is_symlink():
            raise SystemExit(f'missing offline tool: {path}')

    decoder_root = output / 'raw-decoder-replay-v1'
    decoder_result = run_offline(
        [sys.executable, '-B', str(decoder), *sum((['--profile', p] for p in profiles), []),
         '--output-root', str(decoder_root)],
        packet, logs, 'decoder')
    if decoder_result['exit_code'] != 0:
        raise SystemExit('offline decoder failed; terminal retained')

    attribution_root = output / 'attribution-replay-v1'
    attribution_result = run_offline(
        [sys.executable, '-B', str(attributor), '--decoder-root', str(decoder_root),
         '--output-root', str(attribution_root), '--config', str(config),
         *sum((['--histogram', h] for h in histograms), [])],
        packet, logs, 'attribution')
    if attribution_result['exit_code'] != 0:
        raise SystemExit('offline attribution failed; terminal retained')

    # Reconciliation is run against fresh decoder aggregates.  The packet-
    # shaped staging tree consists of hard links, so it adds no second copy of
    # the large JSONL/TSV outputs and cannot silently fall back to the retained
    # decoder directory.
    reconciliation_packet = output / 'reconciliation-input'
    reconciliation_packet.mkdir()
    hardlink_decoder_tree(decoder_root, reconciliation_packet, LABELS)
    (reconciliation_packet / 'tools').mkdir()
    shutil.copy2(decoder, reconciliation_packet / 'tools' / decoder.name)
    reconciliation_output = output / 'xml-multiwindow-reconciliation-v2.json'
    reconciliation_result = run_offline(
        [sys.executable, '-B', str(reconciler), '--packet', str(reconciliation_packet),
         '--source-root', str(source), '--output', str(reconciliation_output)],
        packet, logs, 'reconciliation')
    if reconciliation_result['exit_code'] != 0:
        raise SystemExit('offline reconciliation failed; terminal retained')

    comparisons: list[dict[str, Any]] = []
    aggregate_comparisons: list[dict[str, Any]] = []
    for label, _child in LABELS:
        comparisons.append(compare_json(
            packet / 'raw-decoder-v1' / label / 'summary.json',
            decoder_root / label / 'summary.json'))
        for aggregate_name in ('size-aggregates.tsv', 'trace-aggregates.jsonl',
                               'trace-size-aggregates.tsv'):
            aggregate_comparisons.append(compare_exact(
                packet / 'raw-decoder-v1' / label / aggregate_name,
                decoder_root / label / aggregate_name))
        comparisons.append(compare_json(
            packet / 'attribution-final-v1' / label / 'summary.json',
            attribution_root / label / 'summary.json'))
    comparisons.append(compare_json(
        packet / 'review' / 'xml-multiwindow-source-profile-reconciliation-v2.json',
        reconciliation_output))
    if not all(item['normalized_equal'] for item in comparisons):
        raise SystemExit('one or more relocated semantic comparisons failed; terminal retained')
    if not all(item['exact_equal'] for item in aggregate_comparisons):
        raise SystemExit('one or more freshly decoded aggregate files differ; terminal retained')

    receipt = {
        'schema': 'litchi.performance.0949.offline-relocated-replay.v1',
        'status': 'pass',
        'started_utc': started,
        'ended_utc': datetime.now(timezone.utc).isoformat(),
        'packet_root': str(packet),
        'source_root': str(source),
        'output_root': str(output),
        'commands': [decoder_result, attribution_result, reconciliation_result],
        'comparisons': comparisons,
        'fresh_decoder_aggregate_comparisons': aggregate_comparisons,
        'reconciliation_input': {
            'path': str(reconciliation_packet),
            'mode': 'fresh decoder output exposed by hard links; no retained decoder rows used',
        },
        'scope': {
            'workload_launched': False,
            'binary_launched': False,
            'profiler_launched': False,
            'build_launched': False,
            'decoder': 'offline retained raw profile decode only',
            'attribution': 'offline retained decoder output and histogram attribution',
            'reconciliation': 'offline retained decoded aggregate/source reconciliation; no temporal or adoption claim',
        },
    }
    receipt_path = output / 'offline-replay-terminal.json'
    receipt_path.write_text(json.dumps(receipt, indent=2, sort_keys=True) + '\n')
    print(json.dumps({'status': 'pass', 'receipt': str(receipt_path), 'comparisons': len(comparisons)}, sort_keys=True))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
