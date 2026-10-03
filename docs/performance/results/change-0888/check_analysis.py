#!/usr/bin/env python3
"""Compare full analysis replay and independently parsed VMA sums."""
import argparse
import gzip
import importlib.util
import json
from pathlib import Path

P = Path(__file__).resolve().parent


def load(name, filename):
    spec = importlib.util.spec_from_file_location(name, P / filename)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    if args.output.exists():
        raise ValueError('output already exists')
    # Unique import names avoid the historical baseline verify.py namespace.
    verifier = load('independent_mapping_verifier_0888', 'verify.py')
    analyzer = load('mapping_analyzer_0888', 'analyze.py')
    packets = []
    for packet in (P, P / 'copy-probe'):
        independent = verifier.verify(packet)
        analysis_path = packet / 'analysis.json.gz'
        with gzip.open(analysis_path, 'rt') as f:
            retained = json.load(f)
        assert analyzer.analyse(packet) == retained, str(packet)
        rows = {r['row']['ordinal']: r for r in retained['rows']}
        count = 0
        for child in independent['children']:
            row = rows[child['row']['ordinal']]
            assert row['row'] == child['row']
            assert row['gnu_time_peak_rss_kib'] == child['gnu_time_peak_rss_kib']
            assert row['snapshot_count'] == child['snapshots']
            m = child['sampled_maxima_separate_not_simultaneous']
            parsed = row['sampled_max']['smaps']
            pairs = [
                (m['rss_kib'], parsed['total']['Rss']),
                (m['anonymous_kib'], parsed['total']['Anonymous']),
                (m['exe_rss_kib'], parsed['by_category']['executable_file']['Rss']),
                (m['exe_anonymous_kib'], parsed['by_category']['executable_file']['Anonymous']),
                (m['heap_rss_kib'], parsed['by_category']['heap']['Rss']),
                (m['stack_rss_kib'], parsed['by_category']['stack']['Rss']),
                (m['file_pmd_mapped_kib'], parsed['total']['FilePmdMapped']),
                (m['anon_huge_pages_kib'], parsed['total']['AnonHugePages']),
            ]
            assert all(a * 1024 == b for a, b in pairs), child['row']
            count += child['snapshots']
        packets.append({'packet': str(packet), 'children': len(rows), 'snapshots': count,
                        'analysis_sha256': verifier.sha(analysis_path), 'plan_sha256': verifier.sha(packet / 'plan.json'),
                        'custody_and_semantics': 'pass', 'full_analysis_recomputed': True,
                        'independent_vma_maxima_match': True})
    result = {'status': 'pass', 'packets': packets,
              'helpers': {name: verifier.sha(P / name) for name in ['verify.py', 'analyze.py', 'check_analysis.py']},
              'latency_or_adoption_claim': None}
    with args.output.open('x') as f:
        json.dump(result, f, indent=2, sort_keys=True)
        f.write('\n')
    print('Full analysis and independent mapping totals match for all 24 children.')


if __name__ == '__main__':
    main()
