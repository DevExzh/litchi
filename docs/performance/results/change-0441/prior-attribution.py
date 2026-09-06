#!/usr/bin/env python3
"""Attribute existing 0440 call chains without altering the sealed source."""
import argparse
import gzip
import hashlib
import json
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parent
SOURCE = ROOT.parent / 'change-0440/profiles/after/record/formal/perf-script.txt.gz'


def derive():
    stored = SOURCE.read_bytes()
    raw = gzip.decompress(stored)
    total = transaction = clones = samples = 0
    for block in re.split(r'(?=^normal \d+ )', raw.decode(), flags=re.M):
        match = re.match(r'normal \d+ [\d.]+:\s+(\d+) cycles:u:', block)
        if not match:
            continue
        weight = int(match[1])
        samples += 1
        total += weight
        if re.search(r'\btransaction\+', block):
            transaction += weight
            if 'clone' in block:
                clones += weight
    return {
        'source': str(SOURCE.relative_to(ROOT.parent)),
        'stored_sha256': hashlib.sha256(stored).hexdigest(),
        'raw_sha256': hashlib.sha256(raw).hexdigest(), 'samples': samples,
        'sampled_period_sums': {'total': total, 'transaction': transaction, 'visible_clone_under_transaction': clones},
        'sampled_period_percent': {'transaction': transaction * 100 / total, 'visible_clone_under_transaction': clones * 100 / total},
        'scope': 'Observed call chains include warmups and whole-process setup/oracles. Symbolization and inlining limit visible clone attribution; these are not operation phase timings or allocation fractions.',
    }


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--check', action='store_true')
    args = parser.parse_args()
    result = derive()
    output = ROOT / 'prior-attribution.json'
    if args.check:
        assert json.loads(output.read_text()) == result
    else:
        output.write_text(json.dumps(result, indent=2) + '\n')
    print(json.dumps(result))
