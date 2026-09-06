#!/usr/bin/env python3
"""Group retained 0441 transaction call chains using explicit type markers."""
import argparse
import collections
import gzip
import hashlib
import json
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parent
SOURCE = ROOT.parent / 'change-0441/profiles/after/record/formal/perf-script.txt.gz'
MARKERS = {
    'source_fragments': 'core::option::Option<litchi_odp::codec::content_source::ContentSource>',
    'settings': 'litchi_odp::model::settings::Settings',
    'declarations': 'litchi_odp::model::declaration::Collection',
    'page_metadata': 'litchi_odp::model::page_metadata::Collection',
}


def derive():
    stored = SOURCE.read_bytes(); raw = gzip.decompress(stored)
    counts = collections.Counter(); periods = collections.Counter()
    total = samples = 0
    for block in re.split(r'(?=^normal \d+ )', raw.decode(), flags=re.M):
        match = re.match(r'normal \d+ [\d.]+:\s+(\d+) cycles:u:', block)
        if not match:
            continue
        samples += 1; weight = int(match[1]); total += weight
        if re.search(r'\btransaction\+', block):
            keys = [key for key, marker in MARKERS.items() if marker in block]
            category = keys[0] if len(keys) == 1 else 'unclassified_transaction'
            counts[category] += 1; periods[category] += weight
    return {
        'source': str(SOURCE.relative_to(ROOT.parent)),
        'stored_sha256': hashlib.sha256(stored).hexdigest(),
        'raw_sha256': hashlib.sha256(raw).hexdigest(), 'samples': samples,
        'total_sampled_periods': total, 'markers': MARKERS,
        'sample_counts': dict(counts), 'sampled_period_sums': dict(periods),
        'sampled_period_percent': {key:value*100/total for key,value in periods.items()},
        'scope':'Exclusive observed transaction call-chain groups including warmups. Unmatched/ambiguous frames remain unclassified; incomplete symbolization limits attribution. These are not phase wall times or allocation fractions.',
    }


if __name__ == '__main__':
    parser = argparse.ArgumentParser(); parser.add_argument('--check', action='store_true'); args = parser.parse_args()
    result = derive(); output = ROOT / 'prior-attribution.json'
    if args.check:
        assert json.loads(output.read_text()) == result
    else:
        output.write_text(json.dumps(result, indent=2) + '\n')
    print(json.dumps(result))
