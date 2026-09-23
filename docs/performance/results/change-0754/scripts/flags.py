#!/usr/bin/env python3
"""List every paired process comparison whose after/before ratio moves more
than 5% at p50, p95 or mean, in either direction (0750's method)."""
import json, sys
s = json.load(open(sys.argv[1]))
flags = []
for label, v in sorted(s.items()):
    by = {(p['round'], p['slot']): p for p in v['per_process']}
    for rnd in sorted({p['round'] for p in v['per_process']}):
        for a_slot, b_slot in ((1, 2), (4, 3)):
            a, b = by[(rnd, a_slot)], by[(rnd, b_slot)]
            for metric in ('p50', 'p95', 'mean'):
                ratio = b[metric] / a[metric]
                if abs(ratio - 1) > 0.05:
                    flags.append({'case': label, 'round': rnd, 'pair': f'{a_slot}-{b_slot}', 'metric': metric,
                                  'before_ms': round(a[metric] / 1e6, 4), 'after_ms': round(b[metric] / 1e6, 4),
                                  'change_pct': round((ratio - 1) * 100, 2)})
json.dump(flags, open(sys.argv[2], 'w'), indent=2)
adverse = [f for f in flags if f['change_pct'] > 0]
print(len(flags), 'flags;', len(adverse), 'adverse')
for f in adverse:
    print(f)
