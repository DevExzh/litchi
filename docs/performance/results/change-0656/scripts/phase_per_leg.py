"""Per-leg phase medians (us), run order A1 B1 B2 A2 (change 0656)."""
import json, sys, statistics
S = sys.argv[1]
print("change 0656 - per-leg phase medians (us), run order A1 B1 B2 A2 = before after after before")
legs = {}
for tag in ('A1', 'B1', 'B2', 'A2'):
    d = json.load(open(f'{S}/timing-{tag}.json'))
    legs[tag] = {r['case']: (r['source'].get('pptx_cross_copy') or r['source'].get('pptx_cross_copy_lifecycle') or {}) for r in d['results']}
cases = list(legs['A1'].keys())
for case in cases:
    print(f'\n### {case}')
    keys = [k for k in legs['A1'][case] if k.endswith('_ns') and isinstance(legs['A1'][case][k], list)]
    for k in keys:
        row = ''.join(f'  {tag} {statistics.median(legs[tag][case][k]) / 1000.0:11.3f}' for tag in ('A1', 'B1', 'B2', 'A2'))
        print(f'  {k:16s}{row}')
