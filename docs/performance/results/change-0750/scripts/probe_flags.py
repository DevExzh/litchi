#!/usr/bin/env python3
"""Change 0750: every paired audit-probe comparison moving more than 5% at
the process median, p95 or mean, in either direction."""
import json, sys
s = json.load(open(sys.argv[1]))
flags = []
for case, v in sorted(s.items()):
    by = {(p['round'], p['slot']): p for p in v['per_process']}
    for rnd in sorted({p['round'] for p in v['per_process']}):
        for a_slot, b_slot in ((1, 2), (4, 3)):
            a, b = by[(rnd, a_slot)], by[(rnd, b_slot)]
            for metric in ('median', 'p95', 'mean'):
                ratio = b[metric] / a[metric]
                if abs(ratio - 1) > 0.05:
                    flags.append({'case': case, 'round': rnd, 'pair': f'{a_slot}-{b_slot}', 'metric': metric,
                                  'before_us': round(a[metric] / 1e3, 2), 'after_us': round(b[metric] / 1e3, 2),
                                  'change_pct': round((ratio - 1) * 100, 2)})
json.dump(flags, open(sys.argv[2], 'w'), indent=2)
print(len(flags), 'flags;', sum(1 for f in flags if f['change_pct'] > 0), 'adverse')
