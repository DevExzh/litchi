#!/usr/bin/env python3
"""Independent heaptrack histogram/footer/folded/raw-decoder conservation audit.

This is diagnostic bookkeeping only. It does not classify allocation ownership,
compare latency, or make an adoption claim.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
from pathlib import Path
from typing import Any

BASE = Path('/home/zhuhe/litchi-performance-0949')
LABELS = {
    'h00-baseline-primary-percent-medium': 'baseline-medium',
    'h01-candidate-primary-percent-medium': 'candidate-medium',
    'h02-baseline-primary-percent-dense-sparse': 'baseline-dense',
    'h03-candidate-primary-percent-dense-sparse': 'candidate-dense',
}


class AuditError(RuntimeError):
    pass


def require(condition: Any, message: str) -> None:
    if not condition:
        raise AuditError(message)


def digest(path: Path) -> dict[str, Any]:
    data = path.read_bytes()
    return {'path': str(path), 'bytes': len(data), 'sha256': hashlib.sha256(data).hexdigest()}


def read_json(path: Path) -> Any:
    return json.loads(path.read_text())


def parse_hist(path: Path) -> dict[int, int]:
    result: dict[int, int] = {}
    for lineno, line in enumerate(path.read_text().splitlines(), 1):
        fields = line.split('\t')
        require(len(fields) == 2, f'{path}:{lineno}: malformed histogram row')
        size, calls = map(int, fields)
        require(size >= 0 and calls >= 0, f'{path}:{lineno}: negative histogram value')
        require(size not in result, f'{path}:{lineno}: duplicate histogram size')
        result[size] = calls
    require(result, f'{path}: empty histogram')
    return result


def parse_size_tsv(path: Path) -> dict[int, int]:
    lines = path.read_text().splitlines()
    require(lines and lines[0] == 'size_bytes\tcalls\trequested_bytes', f'{path}: header mismatch')
    result: dict[int, int] = {}
    for lineno, line in enumerate(lines[1:], 2):
        fields = line.split('\t')
        require(len(fields) == 3, f'{path}:{lineno}: malformed size aggregate row')
        size, calls, requested = map(int, fields)
        require(requested == size * calls, f'{path}:{lineno}: requested byte arithmetic mismatch')
        require(size not in result, f'{path}:{lineno}: duplicate decoder size')
        result[size] = calls
    require(result, f'{path}: empty decoder size aggregate')
    return result


def parse_folded(path: Path) -> tuple[int, int]:
    rows = 0
    calls = 0
    for lineno, line in enumerate(path.read_text().splitlines(), 1):
        require(' ' in line and line.rsplit(' ', 1)[1].isdigit(), f'{path}:{lineno}: malformed folded row')
        count = int(line.rsplit(' ', 1)[1])
        require(count >= 0, f'{path}:{lineno}: negative folded count')
        rows += 1
        calls += count
    require(rows, f'{path}: empty folded stack output')
    return rows, calls


def footer_metric(text: str, pattern: str, label: str) -> int:
    match = re.search(pattern, text, re.MULTILINE)
    require(match is not None, f'{label}: footer metric missing')
    return int(match.group(1))


def index_map(index_path: Path) -> dict[str, dict[str, Any]]:
    index = read_json(index_path)
    require(index.get('schema') == 'litchi.performance.0949.root-capture-terminal.v1', 'capture index schema mismatch')
    rows = index.get('artifacts')
    require(isinstance(rows, list), 'capture index artifacts is not a list')
    result: dict[str, dict[str, Any]] = {}
    for row in rows:
        require(isinstance(row, dict) and isinstance(row.get('path'), str), 'bad capture index row')
        require(row['path'] not in result, f'duplicate capture index path {row["path"]}')
        result[row['path']] = row
    return result


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument('--base', type=Path, default=BASE)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    base = args.base.absolute()
    output = args.output.absolute()
    require(not output.exists() and not output.is_symlink(), f'refusing existing output {output}')

    exports_root = base / 'heaptrack-exports-v2'
    exports_path = exports_root / 'exports.json'
    exports = read_json(exports_path)
    require(exports.get('schema') == 'litchi.performance.0949.heaptrack-exports.v1', 'exports schema mismatch')
    rows = exports.get('rows')
    require(isinstance(rows, list) and {r.get('child') for r in rows} == set(LABELS), 'exports child set mismatch')
    capture_index = index_map(base / 'receipts/capture-artifact-index-v1.json')
    observed: list[dict[str, Any]] = []
    for row in sorted(rows, key=lambda x: x['child']):
        child = row['child']
        label = LABELS[child]
        export_dir = exports_root / child
        hist_path = export_dir / 'allocation-size.histogram'
        folded_path = export_dir / 'allocation-stacks.folded'
        stdout_path = export_dir / 'stdout.txt'
        terminal_path = export_dir / 'terminal.json'
        size_path = base / 'raw-decoder-v1' / label / 'size-aggregates.tsv'
        raw_summary_path = base / 'raw-decoder-v1' / label / 'summary.json'
        for path in (hist_path, folded_path, stdout_path, terminal_path, size_path, raw_summary_path):
            require(path.is_file() and not path.is_symlink(), f'missing/nonregular input {path}')
        histogram = parse_hist(hist_path)
        decoder_sizes = parse_size_tsv(size_path)
        require(histogram == decoder_sizes, f'{child}: histogram and raw decoder size maps differ')
        hist_calls = sum(histogram.values())
        hist_bytes = sum(size * count for size, count in histogram.items())
        folded_rows, folded_calls = parse_folded(folded_path)
        footer = stdout_path.read_text()
        footer_calls = footer_metric(footer, r'^calls to allocation functions: (\d+)', f'{child} footer')
        summary = read_json(raw_summary_path)
        require(summary.get('schema') == 'litchi.performance.0949.heaptrack-raw-decode.v1', f'{child}: decoder schema')
        require(summary.get('allocation_events') == hist_calls, f'{child}: raw allocation count mismatch')
        require(summary.get('requested_bytes_total') == hist_bytes, f'{child}: raw requested bytes mismatch')
        require(summary.get('allocation_events') == footer_calls, f'{child}: footer allocation count mismatch')
        require(summary.get('allocation_events') == folded_calls, f'{child}: folded allocation count mismatch')
        require(summary.get('a_plus_adjacency_mismatches') == 0, f'{child}: raw adjacency mismatches')
        require(summary.get('plus_without_metadata') == 0, f'{child}: raw plus metadata mismatches')
        export_terminal = read_json(terminal_path)
        require(export_terminal.get('schema') == 'litchi.performance.0949.heaptrack-export.v1', f'{child}: export terminal schema')
        require(export_terminal.get('exit_code') == 0 and export_terminal.get('error') is None,
                f'{child}: export terminal failed')
        require(export_terminal.get('workload_launched') is False, f'{child}: export launched workload')
        profile = export_terminal.get('profile')
        require(isinstance(profile, dict), f'{child}: export profile descriptor missing')
        profile_rel = f'capture-v1/children/{child}/heaptrack.data.gz.zst'
        indexed = capture_index.get(profile_rel)
        require(indexed is not None, f'{child}: capture profile not indexed')
        require(profile.get('path') == str(base / profile_rel), f'{child}: export profile path mismatch')
        require(profile.get('bytes') == indexed.get('bytes') and profile.get('sha256') == indexed.get('sha256'),
                f'{child}: export profile descriptor differs from sealed capture index')
        observed.append({
            'child': child,
            'label': label,
            'histogram': digest(hist_path),
            'histogram_distinct_sizes': len(histogram),
            'histogram_calls': hist_calls,
            'histogram_requested_bytes': hist_bytes,
            'decoder_size_aggregates': digest(size_path),
            'decoder_summary': digest(raw_summary_path),
            'decoder_allocation_events': summary['allocation_events'],
            'decoder_requested_bytes_total': summary['requested_bytes_total'],
            'decoder_deallocation_events': summary.get('deallocation_events'),
            'decoder_allocation_minus_deallocation': summary.get('allocation_minus_deallocation'),
            'decoder_unknown_any_frame': summary.get('unknown_any_frame'),
            'decoder_unknown_leaf_frame': summary.get('unknown_leaf_frame'),
            'footer': digest(stdout_path),
            'footer_calls': footer_calls,
            'folded': digest(folded_path),
            'folded_distinct_stacks': folded_rows,
            'folded_calls': folded_calls,
            'export_terminal': digest(terminal_path),
            'exact_histogram_map_equal_decoder': True,
            'exact_count_conservation': True,
            'exact_requested_byte_conservation': True,
        })

    result = {
        'schema': 'litchi.performance.0949.independent-heaptrack-conservation.v1',
        'status': 'pass',
        'created_utc': __import__('datetime').datetime.now(__import__('datetime').timezone.utc).isoformat(timespec='microseconds'),
        'workload_launched': False,
        'scope': {
            'raw_decoder': 'Independent exact size-map, count, requested-byte, footer-call, and folded-call conservation. Unknown/deallocation fields are retained per raw decoder summaries.',
            'heaptrack_export': 'The exported heaptrack histogram and allocation flame-stack counts are whole-process requested-allocation diagnostics, not publication-interval attribution.',
            'excluded': 'No timing, ownership, speedup, adoption, or heaptrack publication-only claim is made here.',
        },
        'inputs': {
            'capture_artifact_index': digest(base / 'receipts/capture-artifact-index-v1.json'),
            'exports_index': digest(exports_path),
            'decoder_terminal': digest(base / 'raw-decoder-v1/decoder-terminal.json'),
            'export_helper': digest(base / 'tools/export_heaptrack_0949_v2.py'),
            'heaptrack_print': digest(Path('/usr/bin/heaptrack_print')),
        },
        'rows': observed,
        'checks': {
            'four_children': len(observed) == 4,
            'histogram_exact_size_map_to_raw_decoder': True,
            'histogram_calls_to_raw_decoder': True,
            'histogram_requested_bytes_to_raw_decoder': True,
            'footer_calls_to_raw_decoder': True,
            'folded_calls_to_raw_decoder': True,
            'sealed_profile_descriptors_to_capture_index': True,
            'unknown_and_deallocation_retained': True,
        },
    }
    data = (json.dumps(result, indent=2, sort_keys=True) + '\n').encode()
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open('xb') as stream:
        stream.write(data)
        stream.flush()
        os.fsync(stream.fileno())
    print(json.dumps({'output': str(output), 'bytes': len(data), 'sha256': hashlib.sha256(data).hexdigest()}, sort_keys=True))
    return 0


if __name__ == '__main__':
    try:
        raise SystemExit(main())
    except (AuditError, OSError, KeyError, TypeError, ValueError) as error:
        print(f'0949 independent heaptrack conservation failed: {error}')
        raise SystemExit(1)
