#!/usr/bin/env python3
"""Summarize the per-iteration counter campaign of change 0752.

Each input JSON is one perf_delta.py measurement (a low- and a high-sample
run of one binary). Pairs within each round: slot 1 (before) with slot 2
(after), slot 4 (before) with slot 3 (after).
"""
import glob, json, os, re, statistics, sys

def main():
    directory = sys.argv[1]
    rows = {}
    for path in sorted(glob.glob(os.path.join(directory, '*.json'))):
        m = re.match(r'(.+)-r(\d+)-s(\d)-(before|after)\.json$', os.path.basename(path))
        if not m:
            continue
        d = json.load(open(path))
        rows.setdefault(m.group(1), []).append({'round': int(m.group(2)), 'slot': int(m.group(3)), 'leg': m.group(4),
                                                'instructions': d['per_iteration']['instructions'],
                                                'cycles': d['per_iteration']['cycles'],
                                                'output_sha256': d['output_sha256']})
    summary = []
    print(f"{'case':30} {'instr before':>14} {'instr after':>14} {'change':>8} {'cycles before':>14} {'cycles after':>14} {'change':>8}")
    for name, rs in sorted(rows.items()):
        entry = {'case': name}
        for metric in ('instructions', 'cycles'):
            before = [r[metric] for r in rs if r['leg'] == 'before']
            after = [r[metric] for r in rs if r['leg'] == 'after']
            pairs = []
            for rnd in sorted({r['round'] for r in rs}):
                get = {r['slot']: r for r in rs if r['round'] == rnd}
                for b, a in ((1, 2), (4, 3)):
                    if b in get and a in get:
                        pairs.append(get[a][metric] / get[b][metric] - 1)
            entry[metric] = {'before_median': statistics.median(before), 'after_median': statistics.median(after),
                             'before_all': before, 'after_all': after,
                             'paired_changes': [round(p, 4) for p in pairs],
                             'median_paired_change': round(statistics.median(pairs), 4)}
        entry['output_sha256'] = sorted({s for r in rs for s in r['output_sha256']})
        summary.append(entry)
        i, c = entry['instructions'], entry['cycles']
        print(f"{name:30} {i['before_median']/1e6:13.2f}M {i['after_median']/1e6:13.2f}M {i['median_paired_change']:+8.2%} "
              f"{c['before_median']/1e6:13.2f}M {c['after_median']/1e6:13.2f}M {c['median_paired_change']:+8.2%}")
    json.dump(summary, open(os.path.join(directory, 'summary.json'), 'w'), indent=1)

if __name__ == '__main__':
    main()
