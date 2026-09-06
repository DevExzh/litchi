#!/usr/bin/env python3
"""Interpret unqualified DWARF iteration names using the unique bound binary symbol."""
import argparse
from collections import Counter
import gzip
import json
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile

ROOT = Path(__file__).resolve().parent


def read(path):
    return path.read_text() if path.exists() else gzip.decompress(Path(str(path) + '.gz').read_bytes()).decode()


def derive():
    results = []
    for path in json.loads((ROOT / 'profile-index.json').read_text()):
        receipt = json.loads((ROOT / path).read_text())
        text = read(ROOT / 'profiles' / (receipt['name'] + '-script.log'))
        categories = Counter()
        leaves = {key: Counter() for key in ['iteration', 'corpus_setup', 'other_or_unresolved']}
        api_ancestry = Counter()
        for block in re.split(r'\n\s*\n', text):
            if 'cycles:u:' not in block:
                continue
            symbols = []
            for line in block.splitlines():
                match = re.match(r'\s*[0-9a-f]+\s+(.+?)\s+\([^\n]*\)\s*$', line)
                if match:
                    symbols.append(match.group(1))
            stack = '\n'.join(symbols)
            if any(re.search(r'(?:^|::)run_lifecycle_iteration(?:[<+]|$)', symbol) for symbol in symbols):
                category = 'iteration'
                for name in ['plan_cross_slide_copy', 'publish_cross_slide_copy_to_stream',
                             'from_read_at_with_limits_and_cache_limits_and_execution_context',
                             'try_cache_diagnostics', 'sha256_hex', 'rss_point']:
                    if name in stack:
                        api_ancestry[name] += 1
            elif 'build_pptx_source_backed_cross_copy_corpus' in stack:
                category = 'corpus_setup'
            else:
                category = 'other_or_unresolved'
            categories[category] += 1
            leaves[category][symbols[0] if symbols else '[no decoded frame]'] += 1
        total = sum(categories.values())
        assert total > 0, 'no cycle samples parsed'
        results.append({'name': receipt['name'], 'samples': total,
                        'exclusive_ancestry_counts': dict(categories),
                        'exclusive_ancestry_percent': {key: count * 100 / total for key, count in categories.items()},
                        'iteration_api_ancestry_counts_nonexclusive': dict(api_ancestry),
                        'top_leaf_symbols_by_ancestry': {key: counts.most_common(20) for key, counts in leaves.items()}})
    return {'status': 'pass', 'profiles': results,
            'scope': 'Supplemental analysis of unchanged full-process perf script. The frozen namespace-only analysis remains retained separately. nm -C of the hash-bound binary identifies exactly one run_lifecycle_iteration symbol in pptx_provider_lifecycle; this analysis also recognizes the exact unqualified DWARF name. Unweighted cycles:u sample counts. Explicit non-inlined iteration ancestry separates iterations from corpus setup when frames resolve. Iterations include source construction, observers, checks and drops; API ancestry is nonexclusive. Missing/truncated/unresolved stacks remain other_or_unresolved. Percentages are sampled CPU scope, not API wall-clock fractions or speedup.',
            'performance_claim': None}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--check', action='store_true')
    parser.add_argument('--portable-check', action='store_true')
    args = parser.parse_args()
    result = derive()
    target = ROOT / 'profile-symbol-summary.json'
    if args.check or args.portable_check:
        # Normalize tuple pairs as JSON arrays before comparing.
        assert json.loads(target.read_text()) == json.loads(json.dumps(result))
    else:
        assert not target.exists()
        target.write_text(json.dumps(result, indent=2) + '\n')
    if args.portable_check:
        with tempfile.TemporaryDirectory(prefix='litchi-0429-symbol-replay-') as directory:
            exported = Path(directory) / 'bundle'
            shutil.copytree(ROOT, exported)
            subprocess.run([sys.executable, '-B', str(exported / Path(__file__).name), '--check'], check=True)
    print(json.dumps({'status': 'pass', 'profiles': len(result['profiles']), 'portable': args.portable_check}))


if __name__ == '__main__':
    main()
