#!/usr/bin/env python3
"""Summarize whole-child perf counters without attributing them to edit phases."""
import csv, hashlib, json
from pathlib import Path
HERE = Path(__file__).resolve().parent
profiles = {}
for path in sorted((HERE / 'profiles').glob('*/*.perf.csv')):
    rows = csv.reader(line for line in path.read_text().splitlines() if line and not line.startswith('#'))
    profiles[str(path.relative_to(HERE))] = {'sha256': hashlib.sha256(path.read_bytes()).hexdigest(), 'events': {row[2]: int(row[0]) for row in rows}}
comparisons = []
for count in [1, 32]:
    for kind, left_phase, left_mode, right_mode in [('scalar_before_after', 'before', 'repeated', 'repeated'), ('after_batch_vs_scalar', 'after', 'repeated', 'batch')]:
        left = profiles[f'profiles/{left_phase}/p512-k{count}-owned-{left_mode}.perf.csv']['events']
        right = profiles[f'profiles/after/p512-k{count}-owned-{right_mode}.perf.csv']['events']
        comparisons.append({'replacements': count, 'kind': kind, 'events': {key: {'before': value, 'after': right[key], 'delta_pct': (right[key] / value - 1) * 100 if value else None} for key, value in left.items()}})
result = {'scope': 'Separate whole-child measurements include fixture setup, managed and unmanaged route-specific preflights, and output verification; no operation-local CPU attribution or isolation claim.', 'profiles': profiles, 'comparisons': comparisons}
(HERE / 'profile-summary.json').write_text(json.dumps(result, indent=2) + '\n')
for c in comparisons:
    print(c['kind'], c['replacements'], {key: round(c['events'][key]['delta_pct'], 2) for key in ['cycles', 'instructions', 'branches', 'cache-misses']})
