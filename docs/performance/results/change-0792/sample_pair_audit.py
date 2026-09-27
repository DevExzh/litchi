"""Post-capture exact sample-pair check, separate from the frozen median guard."""
import json
from pathlib import Path
P = Path(__file__).resolve().parent
read = lambda name: json.loads((P/name).read_text())
def guarded_values(sample):
    value = sample['allocation']
    return [value['allocation_calls'], value['allocated_bytes'],
            value['live_bytes_after'] - value['live_bytes_before'],
            value['region_peak_live_bytes'] - value['live_bytes_before']]
pairs = 0
for case in read('plan.json')['cases']:
    for block in range(2):
        stem = f"allocation/{block}-{case['shape']}-{case['mode']}"
        before = read(stem+'-before.json')['samples']
        after = read(stem+'-after.json')['samples']
        assert len(before) == len(after) == 3
        for left, right in zip(before, after):
            assert guarded_values(left) == guarded_values(right), stem
            pairs += 1
assert pairs == 90
print('Exact guarded allocation metrics match in all 90 sample pairs')
