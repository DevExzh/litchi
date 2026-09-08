#!/usr/bin/env python3
"""Replay the targeted ABBA follow-up to larger full-guard latency flags."""
import json
import math
from pathlib import Path
import analyze
import verify

ROOT = Path(__file__).resolve().parent


def require(condition):
    if not condition:
        raise verify.VerificationError("targeted guard custody or protocol differs")


def evaluate(root=ROOT):
    root = root.resolve()
    protocol = verify.read_json(root / 'review-protocol.json', 'review protocol')
    lanes = ['review-A1', 'review-B1', 'review-B2', 'review-A2']
    cases = ['cfb_read_one', 'cfb_shared_concurrent_reads', 'opc_mutated_save', 'xls_fresh_write_to']
    require(protocol['order'] == lanes and protocol['cases'] == cases)
    require((protocol['samples'], protocol['warmups'], protocol['cpu'], protocol['workers'], protocol['expected_rows']) == (100, 5, 2, 1, 7))
    require(protocol['capture_driver_sha256'] == verify.sha256_file(root / 'review_capture.py')[0])
    require(protocol['original_full_comparison_sha256'] == verify.sha256_file(root / 'summary.json')[0])
    bindings = {role: verify.verify_role_binding(root, role) for role in ('control', 'candidate')}
    reports, receipts = [], []
    previous_finish = max(verify._timestamp(verify.read_json(root / lane / 'heaptrack-print.json', 'heap export')['finished_utc'], lane) for lane in ('A-heap', 'B-heap'))
    for lane, role in zip(lanes, ('control', 'candidate', 'candidate', 'control')):
        report = verify.read_json(root / lane / 'report.json', lane)
        receipt = verify.read_json(root / lane / 'receipt.json', lane)
        started = verify.read_json(root / lane / 'started.json', lane)
        binding = bindings[role]
        require(receipt['schema'] == 'litchi-0469-capture-v1')
        require(receipt['lane'] == lane and receipt['role'] == role)
        for field in ('revision', 'binary_sha256'):
            require(receipt[field] == binding[field])
        require(receipt['binding_sha256'] == verify.sha256_file(root / (role + '-binding.json'))[0])
        require(receipt['driver_sha256'] == protocol['capture_driver_sha256'])
        require(receipt['protocol_sha256'] == verify.sha256_file(root / 'review-protocol.json')[0])
        require(receipt['exit_code'] == 0 and receipt['samples'] == 100 and receipt['warmups'] == 5)
        for field in ('clean_before', 'clean_after', 'binary_unchanged', 'report_metadata_matches_clean_role'):
            require(receipt[field] is True)
        require(receipt['cwd'] == binding['build_path'])
        require(all(receipt[key] == value for key, value in started.items()))
        argv = receipt['argv']
        require(argv[:6] == ['taskset', '-c', '2', '/usr/bin/time', '-v', '-o'])
        require(argv[7] == binding['binary_path'] and 'heaptrack' not in argv)
        for flag, value in (('--workers', '1'), ('--warmup', '5'), ('--samples', '100'), ('--case', ','.join(cases)), ('--shape', 'few-large'), ('--writer-shape', 'large')):
            require(argv.count(flag) == 1 and argv[argv.index(flag) + 1] == value)
        artifacts = receipt['artifacts']
        require(set(artifacts) == {'report.json', 'corpus-catalog.json', 'resource.log', 'stdout.log', 'stderr.log', 'started.json'})
        for name, metadata in artifacts.items():
            path = verify.bundle_file(root, f'{lane}/{name}', lane)
            require(verify.sha256_file(path) == (metadata['sha256'], metadata['bytes']))
        start = verify._timestamp(receipt['started_utc'], lane)
        finish = verify._timestamp(receipt['finished_utc'], lane)
        require(previous_finish <= start <= finish)
        previous_finish = finish
        verify.verify_report_metadata(report, lane, binding, 100, 5, cases=cases)
        require(len(report['results']) == 7)
        require(report['configuration']['corpus_shapes'] == ['few-large'])
        require(report['configuration']['writer_shapes'] == ['large'])
        reports.append(report)
        receipts.append(analyze.file_binding(root / lane / 'receipt.json', root))
    abba, tool_path = analyze.load_abba(root)
    try:
        summary = abba.summarize_reports(reports=reports, profile='current-v1')
        status, error = 'valid', None
    except abba.AbbaSummaryInputError as failure:
        # Retain the strict rejection. Never erase varying source counters to
        # manufacture an ABBA acceptance or discard a mismatching result row.
        summary, status, error = None, 'rejected', str(failure)
    comparator, _ = analyze.load_comparator(root)
    observations = []
    for index, first in enumerate(reports[0]['results']):
        item = dict(case=first['case'], corpus=first['corpus'], legs_ns={})
        for lane, report in zip(lanes, reports):
            row = report['results'][index]
            require(row['case'] == first['case'] and row['corpus'] == first['corpus'])
            stats = comparator._latencies(row, lane, 100)
            stats['mean'] = math.fsum(row['elapsed_ns']['samples']) / 100
            item['legs_ns'][lane] = stats
        observations.append(item)
    full = [verify.read_json(root / lane / 'report.json', lane) for lane in ('A-full', 'B-full')]
    full_flags = []
    for before, after in zip(full[0]['results'], full[1]['results']):
        require(before['case'] == after['case'] and before['corpus'] == after['corpus'])
        left = comparator._latencies(before, 'full control', 15)
        right = comparator._latencies(after, 'full candidate', 15)
        left['mean'] = math.fsum(before['elapsed_ns']['samples']) / 15
        right['mean'] = math.fsum(after['elapsed_ns']['samples']) / 15
        for metric in ('p50', 'mean', 'p95', 'p99'):
            delta = 100 * (right[metric] / left[metric] - 1)
            if delta > 5:
                full_flags.append(dict(case=before['case'], corpus=before['corpus'], metric=metric,
                                       baseline_ns=left[metric], candidate_ns=right[metric], delta_percent=delta))
    return dict(schema='litchi-0469-review-v1', scope='Seven targeted guard rows, 100 samples/five warmups; original full-guard flags retained; no registered latency claim',
                abba=summary, strict_abba_status=status, strict_abba_error=error,
                observations=observations, observation_scope='Raw-sample descriptive statistics only; strict ABBA rejection does not authorize comparison claims',
                full_guard_all_five_percent_flags=full_flags, receipts=receipts,
                inputs=[analyze.file_binding(root / lane / 'report.json', root) for lane in lanes],
                protocol=analyze.file_binding(root / 'review-protocol.json', root),
                tool=analyze.file_binding(tool_path, root),
                process_rss={lane: analyze.parse_gnu_rss(root / lane / 'resource.log', root=root, latency_comparison='descriptive normal process') for lane in lanes})


if __name__ == '__main__':
    result = evaluate()
    (ROOT / 'review-summary.json').write_text(json.dumps(result, indent=2, sort_keys=True) + '\n')
    print('targeted guard evidence replayed; strict ABBA:', result['strict_abba_status'])
