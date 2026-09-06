#!/usr/bin/env python3
"""Compare warning-denied harness diagnostics with the retained 0429 baseline."""
import argparse
from collections import Counter
import gzip
import hashlib
import json
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parent


def raw(name):
    path = ROOT / name
    return path.read_bytes() if path.exists() else gzip.decompress(Path(str(path) + '.gz').read_bytes())


def findings(data):
    return Counter(re.findall(r'^error: (.+)\n\s*--> ([^\n]+?):\d+:\d+', data.decode(), re.MULTILINE))


def derive():
    baseline = raw('checks/baseline-harness-strict.log')
    current = raw('checks/harness-strict-final.log')
    before, after = findings(baseline), findings(current)
    changed_region = sum(1 for file, line in re.findall(r'^\s*--> ([^\n]+?):(\d+):\d+', current.decode(), re.MULTILINE)
                         if file == 'src/lib.rs' and (55780 <= int(line) < 55852 or 55948 <= int(line) < 56059 or 59478 <= int(line) < 59592))
    assert before, 'baseline diagnostic extraction is empty'
    assert before == after, {'added': list((after - before).items()), 'removed': list((before - after).items())}
    assert changed_region == 0
    return {'status': 'pass', 'baseline_origin': 'change-0429/checks/harness-strict-final.log.gz',
            'baseline_original_sha256': hashlib.sha256(baseline).hexdigest(),
            'current_log': 'checks/harness-strict-final.log',
            'current_original_sha256': hashlib.sha256(current).hexdigest(),
            'unique_findings': len(after), 'same_message_and_source_file_multiset': True,
            'changed_streaming_region_findings': changed_region,
            'findings': [{'message': message, 'file': path, 'count': count}
                         for (message, path), count in sorted(after.items())]}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--check', action='store_true')
    args = parser.parse_args()
    result = derive()
    target = ROOT / 'checks/strict-debt-comparison.json'
    if args.check:
        assert json.loads(target.read_text()) == result
    else:
        assert not target.exists()
        target.write_text(json.dumps(result, indent=2) + '\n')
    print(json.dumps({'status': 'pass', 'unique_findings': result['unique_findings']}))


if __name__ == '__main__':
    main()
