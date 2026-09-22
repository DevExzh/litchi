#!/usr/bin/env python3
"""Compare shared statistics from separately implemented packet readers."""
import json
from pathlib import Path

P = Path(__file__).resolve().parent
a = json.loads((P / 'order-analysis.json').read_text())
b = json.loads((P / 'independent/independent-analysis.json').read_text())
checked = 0


def compare(left, right):
    global checked
    for key, value in left.items():
        other = right['n' if key == 'count' else key]
        if isinstance(value, dict):
            compare(value, other)
        else:
            assert abs(value - other) < 1e-10, (key, value, other)
            checked += 1


compare(a['groups'], b['window_groups'])
compare(a['within_process_drift'], b['within_process_drift'])
assert len(a['processes']) == len(b['processes']) == 36
assert len(a['pairs']) == len(b['pairs']) == 18
result = dict(status='passed', shared_scalar_statistics=checked, tolerance=1e-10,
              processes=36, samples=1800, pairs=18)
(P / 'independent-comparison.json').write_text(json.dumps(result, indent=2) + '\n')
print(json.dumps(result))
