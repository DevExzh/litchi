#!/usr/bin/env python3
"""Attribute retained 0439 call chains without rerunning or modifying them."""
import collections
import gzip
import hashlib
import json
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parent
SOURCE = ROOT.parent / 'change-0439/profiles/after/record/formal/perf-script.txt.gz'


def derive():
    stored = SOURCE.read_bytes()
    raw = gzip.decompress(stored).decode()
    counts = collections.Counter()
    periods = collections.Counter()
    for block in re.split(r'(?=^normal \d+ )', raw, flags=re.M):
        match = re.match(r'normal \d+ [\d.]+:\s+(\d+) cycles:u:', block)
        if not match:
            continue
        frames = '\n'.join(line for line in block.splitlines() if line.startswith('\t'))
        if 'run_odp_existing_append_lifecycle+' not in frames:
            category = 'outside_runner'
        elif re.search(r'\bcommit\+', frames):
            category = 'commit_stack'
        elif re.search(r'\btransaction\+', frames):
            category = 'transaction_stack'
        elif re.search(r'\bfrom_bytes\+', frames):
            category = 'open_stack'
        else:
            category = 'runner_other'
        counts[category] += 1
        periods[category] += int(match.group(1))
    assert sum(counts.values()) == 6850
    return {
        'source': str(SOURCE.relative_to(ROOT.parent)),
        'stored_sha256': hashlib.sha256(stored).hexdigest(),
        'raw_sha256': hashlib.sha256(raw.encode()).hexdigest(),
        'sample_count': sum(counts.values()),
        'sample_counts': dict(counts),
        'sampled_period_sums': dict(periods),
        'sampled_period_percent': {key: value * 100 / sum(periods.values()) for key, value in periods.items()},
        'scope': 'Exclusive observed call-chain groups within the existing-append runner, including warmups. Other runner stacks include untimed checks/destruction. Outside-runner samples include fixture setup. Missing or ambiguous frames remain outside named groups. These are sampled CPU attribution, not phase wall times or allocations.',
        'diagnostics': 'The retained perf script includes 13 addr2line warnings; no symbolization completeness claim.'
    }


if __name__ == '__main__':
    result = derive()
    (ROOT / 'attribute-stacks.json').write_text(json.dumps(result, indent=2) + '\n')
    print(json.dumps(result, indent=2))
