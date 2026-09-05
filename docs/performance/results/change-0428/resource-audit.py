#!/usr/bin/env python3
"""Audit retained resource scopes and RSS repeat flags without the executable."""
import argparse
import gzip
import json
from pathlib import Path
import re
import statistics

ROOT = Path(__file__).resolve().parent


def artifact(path):
    return path.read_bytes() if path.exists() else gzip.decompress(Path(str(path) + '.gz').read_bytes())


def flatten(value, prefix=''):
    result = {}
    for key, item in value.items():
        name = prefix + key
        if isinstance(item, dict):
            result.update(flatten(item, name + '.'))
        elif type(item) is int:
            result[name] = item
    return result


def derive():
    rows = []
    signatures = {}
    final_zero = 0
    compared_points = 0
    for name in json.loads((ROOT / 'capture-index.json').read_text()):
        receipt = json.loads((ROOT / name).read_text())
        report = json.loads((ROOT / 'capture' / (receipt['name'] + '.json')).read_text())
        rss, hwm, entries = [], [], []
        role = (receipt['scenario'], receipt['corpus'])
        for sample in report['samples_raw']:
            signature = []
            entries.append(sample['phases'][0]['rss']['rss_bytes'])
            for phase in sample['phases']:
                values = flatten({key: value for key, value in phase.items() if key != 'rss'})
                signature.append((phase['label'], values))
                rss.append(phase['rss']['rss_bytes'])
                hwm.append(phase['rss']['vm_hwm_bytes'])
                compared_points += 1
            assert signatures.setdefault(role, signature) == signature, 'non-RSS observations differ'
            final = sample['phases'][-1]
            for owner in ['source_budget', 'destination_budget']:
                assert all(final[owner][key] == 0 for key in ['memory_used', 'objects_used', 'depth_used'])
            final_zero += 1
        resource = artifact(ROOT / 'capture' / (receipt['name'] + '-resource.log')).decode()
        maximum = re.search(r'Maximum resident set size \(kbytes\): (\d+)', resource)
        assert maximum
        rows.append({'name': receipt['name'], 'scenario': role[0], 'corpus': role[1],
                     'repeat': receipt['repeat'], 'entry_rss_median': statistics.median(entries),
                     'all_phase_rss_min': min(rss), 'all_phase_rss_max': max(rss),
                     'all_phase_vmhwm_min': min(hwm), 'all_phase_vmhwm_max': max(hwm),
                     'external_maxrss_bytes': int(maximum.group(1)) * 1024,
                     'samples': len(report['samples_raw'])})
    summary = json.loads((ROOT / 'summary.json').read_text())
    triggers = summary['repeat_review_triggers']
    assert all(row['scope'] == 'process RSS point' for row in triggers)
    comparisons = []
    for role in sorted(signatures):
        pair = {row['repeat']: row for row in rows if (row['scenario'], row['corpus']) == role}
        before, after = pair['R1']['entry_rss_median'], pair['R2']['entry_rss_median']
        comparisons.append({'scenario': role[0], 'corpus': role[1], 'r1_entry_rss': before,
                            'r2_entry_rss': after, 'entry_rss_change_percent': 100 * (after - before) / before})
    return {'status': 'pass', 'retained_samples': final_zero, 'phase_points': compared_points,
            'non_rss_numeric_observations_identical_within_and_across_repeats': True,
            'final_releasable_budget_zero_samples': final_zero,
            'rss_repeat_review_triggers': len(triggers), 'rss_processes': rows,
            'entry_rss_comparisons': comparisons,
            'interpretation': 'Descriptive only. Entry RSS already differs; procfs VmHWM and external process maxima include setup/warmups. No causal allocator or RSS-release conclusion.'}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--check', action='store_true')
    args = parser.parse_args()
    result = derive()
    target = ROOT / 'resource-audit.json'
    if args.check:
        assert json.loads(target.read_text()) == result
    else:
        assert not target.exists()
        target.write_text(json.dumps(result, indent=2) + '\n')
    print(json.dumps({'status': 'pass', 'samples': result['retained_samples'],
                      'phase_points': result['phase_points'], 'rss_flags': result['rss_repeat_review_triggers']}))


if __name__ == '__main__':
    main()
