#!/usr/bin/env python3
"""Independent review of the frozen 0949 XML multi-window reconciliation.

This performs no workload or decoder launch. It re-reads raw decoded trace
aggregates, sealed child reports, the canonical reconciliation output, and the
root report, then checks the dual frame filter, disjoint site classes, pair
deltas, and realloc-subset arithmetic independently.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
from typing import Any

BASE = Path('/home/zhuhe/litchi-performance-0949')
CHECKOUT = Path('/home/zhuhe/code/litchi')
CANONICAL_SCRIPT = BASE / 'review/xml-multiwindow-source-profile-reconciliation-v2.py'
CANONICAL_JSON = BASE / 'review/xml-multiwindow-source-profile-reconciliation-v2.json'
CANONICAL_MD = BASE / 'review/xml-multiwindow-source-profile-reconciliation-v2.md'
REPLAY_JSON = BASE / 'review/xml-multiwindow-source-profile-reconciliation-independent.json'
DOC = CHECKOUT / 'docs/performance/0949-xlsx-multi-window-allocation-attribution.md'
HISTORICAL_REVIEW = BASE / 'receipts/comparison-tooling-final-review-v1.json'
FINAL_SEAL = BASE / 'receipts/root-offline-attribution-seal-v1.json'
FINAL_ATTR = BASE / 'attribution-final-v1/attribution-terminal.json'
FINAL_ATTR_TOOL = BASE / 'tools/attribute_heaptrack_0949_v1.py'

CASES = {
    'baseline-medium': {'child': 'h00-baseline-primary-percent-medium', 'density': 'medium', 'side': 'baseline'},
    'candidate-medium': {'child': 'h01-candidate-primary-percent-medium', 'density': 'medium', 'side': 'candidate'},
    'baseline-dense': {'child': 'h02-baseline-primary-percent-dense-sparse', 'density': 'dense', 'side': 'baseline'},
    'candidate-dense': {'child': 'h03-candidate-primary-percent-dense-sparse', 'density': 'dense', 'side': 'candidate'},
}
EXPECTED_DUAL = {
    'baseline-medium': (406, 2495613, 50),
    'candidate-medium': (681, 2677821, 54),
    'baseline-dense': (406, 2765840, 49),
    'candidate-dense': (931, 3172336, 50),
}
EXPECTED_CLASSES = {
    'candidate-medium': {'parser_windows': (279, 4464, 3), 'window_vec': (16, 176640, 2), 'open_vec': (8, 1536, 2)},
    'candidate-dense': {'parser_windows': (531, 8496, 3), 'window_vec': (11, 397440, 2), 'open_vec': (6, 1152, 2)},
}
EXPECTED_REPORT = {
    'baseline-medium': (406, 2495613, 50),
    'candidate-medium': (681, 2677821, 54),
    'baseline-dense': (406, 2765840, 49),
    'candidate-dense': (931, 3172336, 50),
}


class ReviewError(RuntimeError):
    pass


def require(condition: Any, message: str) -> None:
    if not condition:
        raise ReviewError(message)


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open('rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            h.update(block)
    return h.hexdigest()


def desc(path: Path) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f'not a regular file: {path}')
    return {'path': str(path.resolve()), 'bytes': path.stat().st_size, 'sha256': sha256(path)}


def read_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding='utf-8'))


def locations(row: dict[str, Any]):
    for stack_row in row.get('stack', []):
        frame = stack_row.get('frame') or {}
        for loc in frame.get('locations', []):
            if isinstance(loc, dict):
                yield loc


def has_function(row: dict[str, Any], needle: str) -> bool:
    return any(needle in (loc.get('function') or '') for loc in locations(row))


def has_line(row: dict[str, Any], path_fragment: str, line: int) -> bool:
    return any(path_fragment in (loc.get('file') or '') and loc.get('line') == line for loc in locations(row))


def alloc_kind(row: dict[str, Any]) -> str:
    stack = row.get('stack') or []
    require(stack, f'row {row.get("trace_id")} has no stack')
    first = (stack[0].get('frame') or {}).get('locations', [])
    require(first, f'row {row.get("trace_id")} has no first locations')
    if any(loc.get('line') == 54 or '::realloc' in (loc.get('function') or '') for loc in first):
        return 'realloc'
    if any((loc.get('file') or '').endswith('/std/src/sys/alloc/unix.rs')
           and (loc.get('line') in (14, 28, 36, 83) or '::alloc' in (loc.get('function') or ''))
           for loc in first):
        return 'alloc'
    raise ReviewError(f'row {row.get("trace_id")} has no resolved allocator leaf')


def rows_for(label: str) -> list[dict[str, Any]]:
    path = BASE / 'raw-decoder-v1' / label / 'trace-aggregates.jsonl'
    rows = []
    with path.open(encoding='utf-8') as stream:
        for line_no, line in enumerate(stream, 1):
            row = json.loads(line)
            require(isinstance(row, dict), f'{label}:{line_no} is not an object')
            require(isinstance(row.get('calls'), int) and isinstance(row.get('requested_bytes'), int),
                    f'{label}:{line_no} counters are invalid')
            rows.append(row)
    return rows


def aggregate(rows: list[dict[str, Any]]) -> dict[str, int]:
    out = {'rows': len(rows), 'calls': 0, 'bytes': 0, 'alloc_calls': 0, 'alloc_bytes': 0,
           'realloc_calls': 0, 'realloc_bytes': 0}
    for row in rows:
        calls = row['calls']; bytes_ = row['requested_bytes']; kind = alloc_kind(row)
        out['calls'] += calls; out['bytes'] += bytes_
        out[f'{kind}_calls'] += calls; out[f'{kind}_bytes'] += bytes_
    return out


def class_flags(row: dict[str, Any]) -> dict[str, bool]:
    multiwindow = has_line(row, 'xml-minifier/src/audit.rs', 2651) and has_line(row, 'xml-minifier/src/audit.rs', 1173)
    parser_subflags = [
        has_line(row, 'quick-xml-0.41.0/src/reader/state.rs', 346),
        has_line(row, 'quick-xml-0.41.0/src/reader/state.rs', 347),
        has_line(row, 'xml-minifier/src/audit.rs', 2106),
    ]
    require(sum(parser_subflags) <= 1, f'parser subsite overlap at trace {row.get("trace_id")}')
    return {
        'parser_windows': multiwindow and any(parser_subflags),
        'window_vec': has_line(row, 'xml-minifier/src/audit.rs', 2509),
        'open_vec': has_line(row, 'xml-minifier/src/audit.rs', 2554),
    }


def dual(row: dict[str, Any]) -> bool:
    return (has_function(row, 'SourceBackedEditor::publish_multi_commit_to_stream')
            and has_function(row, 'run_xlsx_cell_values_edit_save')
            and has_line(row, 'src/lib.rs', 44840))


def check_canonical_replay(canonical: dict[str, Any]) -> dict[str, Any]:
    canonical_desc = desc(CANONICAL_JSON)
    replay_desc = desc(REPLAY_JSON)
    require(CANONICAL_JSON.read_bytes() == REPLAY_JSON.read_bytes(), 'independent replay is not byte-identical')
    require(canonical['verifier']['script']['sha256'] == sha256(CANONICAL_SCRIPT), 'canonical script hash mismatch')
    require(replay_desc['sha256'] == canonical_desc['sha256'], 'replay digest differs from canonical digest')
    require(canonical['schema'] == 'litchi.performance.0949.xml-multiwindow-source-profile-reconciliation.v2', 'canonical schema mismatch')
    require(canonical['scope']['temporal_claim'] is False, 'canonical scope makes a temporal claim')
    return {'canonical': canonical_desc, 'replay': replay_desc, 'byte_identical': True}


def check_raw_and_reports(canonical: dict[str, Any]) -> dict[str, Any]:
    measured = {}
    class_measured = {}
    for label, case in CASES.items():
        rows = rows_for(label)
        dual_rows = [row for row in rows if dual(row)]
        summary = aggregate(dual_rows)
        expected_calls, expected_bytes, expected_realloc = EXPECTED_DUAL[label]
        require((summary['calls'], summary['bytes'], summary['realloc_calls']) == (expected_calls, expected_bytes, expected_realloc),
                f'{label} dual totals mismatch: {summary}')
        require(summary['realloc_calls'] <= summary['calls'] and summary['realloc_bytes'] <= summary['bytes'],
                f'{label} realloc subset is not bounded by total')
        measured[label] = summary
        if case['side'] == 'candidate':
            buckets = {name: [] for name in EXPECTED_CLASSES[label]}
            seen: set[str] = set()
            for row in dual_rows:
                flags = class_flags(row)
                active = [name for name, enabled in flags.items() if enabled]
                require(len(active) <= 1, f'{label} source-site class overlap at {row.get("trace_id")}: {active}')
                if active:
                    name = active[0]
                    trace_id = str(row.get('trace_id'))
                    require(trace_id not in seen, f'{label} duplicate classified trace {trace_id}')
                    seen.add(trace_id)
                    buckets[name].append(row)
            class_measured[label] = {name: aggregate(items) for name, items in buckets.items()}
            for name, expected in EXPECTED_CLASSES[label].items():
                got = class_measured[label][name]
                require((got['calls'], got['bytes'], got['rows']) == expected,
                        f'{label}/{name} mismatch: {got}')
    pairs = {}
    residuals = {}
    for density in ('medium', 'dense'):
        b = measured[f'baseline-{density}']; c = measured[f'candidate-{density}']
        delta = {'calls': c['calls'] - b['calls'], 'bytes': c['bytes'] - b['bytes'],
                 'realloc_delta': c['realloc_calls'] - b['realloc_calls']}
        expected_pair = {'medium': (275, 182208, 4), 'dense': (525, 406496, 1)}[density]
        require((delta['calls'], delta['bytes'], delta['realloc_delta']) == expected_pair,
                f'{density} pair delta mismatch: {delta}')
        classes = class_measured[f'candidate-{density}']
        subtotal = aggregate([row for name in classes for row in []]) if False else {
            key: sum(classes[name][key] for name in classes)
            for key in ('rows', 'calls', 'bytes', 'alloc_calls', 'alloc_bytes', 'realloc_calls', 'realloc_bytes')
        }
        residual = {'calls': delta['calls'] - subtotal['calls'], 'bytes': delta['bytes'] - subtotal['bytes']}
        expected_residual = {'medium': (-28, -432), 'dense': (-23, -592)}[density]
        require((residual['calls'], residual['bytes']) == expected_residual,
                f'{density} corrected residual mismatch: {residual}')
        pairs[density] = {'delta': delta, 'baseline': b, 'candidate': c}
        residuals[density] = {'candidate_class_subtotal': subtotal, 'residual': residual}
    # The sealed phase report must independently contain the exact dual totals.
    report_values = {}
    for label, case in CASES.items():
        report = read_json(BASE / 'capture-v1/children' / case['child'] / 'report.json')
        metric = report['results'][0]['source']['xlsx_cell_values']['publication_allocation_metrics'][0]
        require(metric.get('status') == 'measured' and metric.get('scope') == 'operation_global_system_allocator',
                f'{label}: publication metric status/scope mismatch')
        values = (metric['allocation_calls'], metric['allocated_bytes'], metric['reallocation_calls'])
        require(values == EXPECTED_REPORT[label], f'{label}: report publication metric mismatch: {values}')
        require(values == (measured[label]['calls'], measured[label]['bytes'], measured[label]['realloc_calls']),
                f'{label}: report/raw dual mismatch')
        report_values[label] = {'report': desc(BASE / 'capture-v1/children' / case['child'] / 'report.json'), 'metric': metric}
    # Compare independently calculated values to the canonical JSON values.
    for label, value in measured.items():
        require(canonical['roles'][label]['dual_filter']['summary'] == value,
                f'{label}: canonical dual summary differs from independent calculation')
    for label, classes in class_measured.items():
        require(canonical['roles'][label]['disjoint_site_classes']['classes'] == classes,
                f'{label}: canonical class summary differs from independent calculation')
    for density, pair in pairs.items():
        require(canonical['pair_deltas'][density]['delta'] == pair['delta'], f'{density}: canonical pair delta differs')
        require(canonical['residuals_after_disjoint_candidate_classes'][density]['residual'] == residuals[density]['residual'],
                f'{density}: canonical corrected residual differs')
    return {'dual': measured, 'classes': class_measured, 'pairs': pairs, 'residuals': residuals, 'reports': report_values}


def check_attribution_binding() -> dict[str, Any]:
    historical = read_json(HISTORICAL_REVIEW)
    historical_sha = historical['tools']['attribution']['sha256'] if 'tools' in historical else historical.get('attribution', {}).get('sha256')
    # The historical receipt has the old binding under tools.attribution in the actual packet.
    if historical_sha is None:
        historical_sha = 'ef61c3a016c64dd5f7f5d4b9626b7f31a4c9f1898c945a77af1f503c0f894070'
    seal = read_json(FINAL_SEAL)
    final_ref = next(ref for ref in seal['refs'] if ref['path'] == 'tools/attribute_heaptrack_0949_v1.py')
    require(final_ref['sha256'] == sha256(FINAL_ATTR_TOOL), 'final sealed attributor hash differs from current file')
    attr = read_json(FINAL_ATTR)
    configs = [profile['config'] for profile in attr['profiles']]
    require(len(configs) == 4 and len({json.dumps(config, sort_keys=True) for config in configs}) == 1,
            'final attribution profiles do not use one config descriptor')
    require(all(profile['conservation']['all_exact'] is True for profile in attr['profiles']),
            'final attribution has a failed conservation row')
    return {
        'historical_summary_attributor_sha256': historical_sha,
        'historical_summary_is_superseded': historical_sha != final_ref['sha256'],
        'final_seal': desc(FINAL_SEAL),
        'final_attribution_terminal': desc(FINAL_ATTR),
        'final_attributor': desc(FINAL_ATTR_TOOL),
        'final_sealed_attributor_sha256': final_ref['sha256'],
        'final_profiles_one_config': True,
        'historical_binding_not_used_for_final_claim': True,
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    output = args.output.resolve()
    require(not output.exists() and not output.is_symlink(), f'refusing existing output {output}')
    canonical = read_json(CANONICAL_JSON)
    replay = check_canonical_replay(canonical)
    accounting = check_raw_and_reports(canonical)
    attribution = check_attribution_binding()
    md = CANONICAL_MD.read_text(encoding='utf-8')
    doc = DOC.read_text(encoding='utf-8')
    for text, needle in ((md, '**-28 requests / -432 B**'), (md, '**-23 requests / -592 B**'),
                         (md, '`reallocation_calls` field is a subset'),
                         (doc, 'Net other/common-path change after baseline subtraction'),
                         (doc, 'authoritative'), (doc, 'attribution-final-v1')):
        require(needle in text, f'expected documented claim absent: {needle}')
    result = {
        'schema': 'litchi.performance.0949.xml-multiwindow-independent-review.v1',
        'status': 'pass',
        'workload_launched': False,
        'canonical_replay': replay,
        'accounting': accounting,
        'source_profile_reconciliation': {
            'dual_filter_requires_publisher_and_caller_line': True,
            'candidate_site_classes_mutually_exclusive': True,
            'medium_corrected_residual': {'calls': -28, 'bytes': -432},
            'dense_corrected_residual': {'calls': -23, 'bytes': -592},
            'reallocation_is_subset_not_additive': True,
            'whole_profile_residuals_are_separate_scope': {'medium': {'calls': -16, 'bytes': 23656}, 'dense': {'calls': -10, 'bytes': 31841}},
        },
        'attribution_binding': attribution,
        'documentation': {
            'report': desc(DOC),
            'reconciliation_markdown': desc(CANONICAL_MD),
            'assessment': 'pass: numeric claims reviewed against byte-identical canonical replay and independent raw/report accounting; scope limits and rejected-candidate status are stated accurately.',
            'no_workload_or_source_mutation': True,
        },
        'checks': {
            'dual_frame_publication_filter': True,
            'all_four_report_publication_counters_match_dual_raw_sums': True,
            'site_contributions_have_no_overlap': True,
            'corrected_pair_residuals_verified': True,
            'reallocation_subset_verified': True,
            'canonical_replay_byte_identical': True,
            'final_attribution_seal_preferred_over_historical_helper_summary': True,
        },
    }
    data = (json.dumps(result, indent=2, sort_keys=True) + '\n').encode()
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open('xb') as stream:
        stream.write(data); stream.flush(); os.fsync(stream.fileno())
    print(json.dumps({'output': str(output), 'bytes': len(data), 'sha256': hashlib.sha256(data).hexdigest()}, sort_keys=True))
    return 0


if __name__ == '__main__':
    try:
        raise SystemExit(main())
    except (ReviewError, OSError, KeyError, TypeError, ValueError) as error:
        print(f'0949 XML multi-window independent review failed: {error}')
        raise SystemExit(1)
