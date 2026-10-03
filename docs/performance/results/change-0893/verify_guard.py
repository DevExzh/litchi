#!/usr/bin/env python3
"""Validate the serialized guard qualification report, never profiler timings."""
import argparse
import copy
import hashlib
import json
from pathlib import Path

P = Path(__file__).resolve().parent
PREFIX = '0893_FULL_CELL_REPORT='
CONTRACT = 'post-endpoint-per-worksheet-tolerance-v1'
ARCHIVE = '55e0901c093c27d80bcd4e5fef888f00798654cdef8c0d79d130a2ce7cbe0ae8'
DIGEST = '8d29b40718aa9d5b08c746f841195bba17ab8e01608b12e6093daf160c00764b'
CENSUS = '552c35c8b8f958de28792f00c1e9e831ae884471aa28b5b2a7a1479bd06c5207'


def require(condition, message):
    if not condition:
        raise ValueError(message)


def identity(path):
    path = path.resolve()
    data = path.read_bytes()
    return {'path': str(path), 'bytes': len(data), 'sha256': hashlib.sha256(data).hexdigest()}


def validate(result, warmups=1, samples=3):
    require(result['case'] == 'xlsx_producer_medium_source_full_cell_scan', 'case')
    corpus = result['corpus']
    require(corpus['archive_sha256'] == ARCHIVE and corpus['archive_bytes'] == 22739, 'corpus')
    require(corpus['xlsx']['sheet_count'] == 4, 'worksheet count')
    stats = result['elapsed_ns']
    require(len(stats['samples']) == samples and all(type(x) is int and x > 0 for x in stats['samples']), 'samples')
    require(sorted(stats['sample_order']) == list(range(samples)), 'sample permutation')
    scan = result['source']['producer_full_cell_scan']
    require(scan.get('guard_contract') == CONTRACT, 'guard contract missing or wrong')
    require(scan['sample_indices'] == stats['sample_order'], 'guard sample alignment')
    require(result['operation_metrics']['sample_indices'] == stats['sample_order'], 'metric alignment')
    require(scan['archive_sha256'] == ARCHIVE and scan['archive_bytes'] == 22739, 'scan archive')
    for field, expected in [('cell', 4096), ('numeric', 3686), ('text', 410)]:
        require(scan[f'expected_{field}_count'] == expected, field + ' expected count')
        require(scan[f'observed_{field}_counts'] == [expected] * samples, field + ' observed count')
    require(scan['oracle_digest_sha256'] == DIGEST and scan['observed_digest_sha256'] == [DIGEST] * samples, 'cell digest')
    require(scan['marker_refusal_census_sha256'] == CENSUS, 'static census')
    require(len(scan['source_version_stable']) == samples and all(v is True for v in scan['source_version_stable']), 'final source fences')
    require(scan.get('forced_worksheet_counts') == [4] * samples, 'measured forced count')
    require(scan.get('forced_worksheet_tolerance_clean') == [True] * samples and all(v is True for v in scan['forced_worksheet_tolerance_clean']), 'measured clean')
    require(scan.get('warmup_forced_worksheet_counts') == [4] * warmups, 'warmup forced count')
    require(scan.get('warmup_forced_worksheet_tolerance_clean') == [True] * warmups and all(v is True for v in scan['warmup_forced_worksheet_tolerance_clean']), 'warmup clean')
    require(scan.get('forced_worksheet_guard_iterations') == warmups + samples, 'guard iterations')
    require(scan['runtime_route_observed'] is False and scan['physical_cold_eligible'] is False, 'scope flags')
    require(all(phrase in scan['timing_scope'] for phrase in ('per-worksheet tolerance guard', 'aggregate tolerance report', 'process/allocation endpoint snapshots', 'post-guard source-freshness fence')), 'guard timing scope')
    return {'warmup_iterations': warmups, 'measured_iterations': samples, 'worksheets_per_iteration': 4, 'guard_contract': CONTRACT}


def negative_controls(result):
    mutations = {
        'old_guard_contract': ('guard_contract', 'aggregate-only'),
        'missing_guard_contract': ('guard_contract', None),
        'short_measured_counts': ('forced_worksheet_counts', [4, 4]),
        'unchecked_sheet': ('forced_worksheet_counts', [4, 3, 4]),
        'dirty_measured_report': ('forced_worksheet_tolerance_clean', [True, False, True]),
        'nonboolean_clean': ('forced_worksheet_tolerance_clean', [1, 1, 1]),
        'missing_warmup': ('warmup_forced_worksheet_counts', []),
        'dirty_warmup_report': ('warmup_forced_worksheet_tolerance_clean', [False]),
        'wrong_guard_iterations': ('forced_worksheet_guard_iterations', 3),
        'unstable_source': ('source_version_stable', [True, False, True]),
        'wrong_digest': ('observed_digest_sha256', [DIGEST, '0' * 64, DIGEST]),
        'invalid_sample_alignment': ('sample_indices', [0, 0, 0]),
    }
    passed = []
    for name, (key, value) in mutations.items():
        altered = copy.deepcopy(result)
        scan = altered['source']['producer_full_cell_scan']
        if value is None:
            scan.pop(key, None)
        else:
            scan[key] = value
        try:
            validate(altered)
        except (ValueError, KeyError):
            passed.append(name)
        else:
            raise ValueError('negative control accepted: ' + name)
    return passed


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--log', required=True, type=Path)
    parser.add_argument('--output-root', required=True, type=Path)
    args = parser.parse_args()
    lines = [line.split(PREFIX, 1)[1] for line in args.log.read_text().splitlines() if PREFIX in line]
    require(len(lines) == 1, 'expected one serialized qualification report')
    result = json.loads(lines[0])
    checked = validate(result)
    negatives = negative_controls(result)
    historical = json.loads((P.parent / 'change-0892/capture/child-00/report.json').read_text())['results'][0]
    try:
        validate(historical, warmups=10, samples=2000)
    except ValueError as error:
        require(str(error) == 'guard contract missing or wrong', 'historical report failed for unrelated reason')
    else:
        raise ValueError('historical aggregate-only report accepted')
    args.output_root.mkdir(parents=True, exist_ok=False)
    report = args.output_root / 'report.json'
    report.write_text(json.dumps(result, indent=2, sort_keys=True) + '\n')
    receipt = {'schema': 'litchi.full-cell-guard-qualification-0893.v1', 'status': 'pass', 'log': identity(args.log), 'verifier': identity(Path(__file__)), 'report': identity(report), 'checked': checked, 'negative_controls_rejected': negatives, 'historical_0892_guard_rejected': True, 'performance_claim': None, 'scope': 'Debug test report semantic/guard evidence; elapsed and process values are not performance measurements.'}
    (args.output_root / 'verification.json').write_text(json.dumps(receipt, indent=2, sort_keys=True) + '\n')
    print(json.dumps({'status': 'pass', 'checked': checked, 'negative_controls': len(negatives), 'historical_guard_rejected': True}))


if __name__ == '__main__':
    main()
