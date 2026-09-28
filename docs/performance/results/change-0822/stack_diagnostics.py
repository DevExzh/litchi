"""Offline stack-depth diagnostics; absence of markers does not prove full unwinding."""
import gzip
import hashlib
import json
from pathlib import Path
import re
import sys
from collections import Counter

P = Path(__file__).resolve().parent
assert sys.argv[1:] in (['--write'], ['--check'])
rows = []
for member in json.loads((P / 'perf/compression.json').read_text()):
    if member['kind'] != 'frames':
        continue
    stored = Path(member['compressed']['path']).read_bytes()
    assert hashlib.sha256(stored).hexdigest() == member['compressed']['sha256']
    raw = gzip.decompress(stored)
    assert hashlib.sha256(raw).hexdigest() == member['original']['sha256']
    blocks = [block.splitlines() for block in raw.decode().split('\n\n') if block.strip()]
    depths = Counter(len(block) - 1 for block in blocks)
    lines = raw.decode().splitlines()
    rows.append({
        'repeat': member['repeat'], 'frames_sha256': member['original']['sha256'],
        'samples': len(blocks), 'max_observed_stack_depth': max(depths),
        'stack_depth_counts': dict(sorted(depths.items())),
        'empty_stack_samples': depths[0],
        'explicit_truncation_lines': [line for line in lines if re.search('truncat', line, re.I)],
        'stacks_ending_unknown': sum(bool(len(block) > 1 and '[unknown]' in block[-1]) for block in blocks),
        'complete_unwinding_proven': False,
        'limitation': 'Stack depth and explicit markers are diagnostics only; missing callers and unmarked truncation cannot be ruled out.',
    })
value = {'schema': 'litchi.performance.0822.stack-diagnostics.v1', 'repeats': rows}
encoded = json.dumps(value, indent=2, sort_keys=True) + '\n'
out = P / 'stack-diagnostics.json'
if sys.argv[1] == '--write':
    assert not out.exists()
    out.write_text(encoded)
else:
    assert out.read_text() == encoded
print('0822 stack diagnostics PASS')
