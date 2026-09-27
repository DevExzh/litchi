"""Independent raw-count audit; no native workload or primary analyzer imports."""
import argparse
from collections import Counter
import hashlib
import json
from pathlib import Path
import re

P = Path(__file__).resolve().parent
OWNER = re.compile(r'^namespace_uri_probe::capture_region_0793::h[0-9a-f]+(?: \(.*\))?$')

def stacks(path):
    rows = Counter()
    for line in path.read_text().splitlines():
        stack, count = line.rsplit(' ', 1)
        assert count.isdecimal() and int(count) > 0
        rows[stack] += int(count)
    return rows

def owns(stack):
    return any(OWNER.fullmatch(frame) for frame in stack.split(';'))

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--write', action='store_true')
    parser.add_argument('--check', action='store_true')
    args = parser.parse_args()
    load = lambda path: json.loads(path.read_text())
    rows = []
    count_reports = count_samples = 0
    for lane in ['controls', 'heaptrack']:
        receipts = load(P / lane / 'receipts.json')
        assert len(receipts) == (40 if lane == 'controls' else 10)
        for receipt in receipts:
            assert receipt['exit_code'] == 0
            report = receipt['report']
            path = P / lane / Path(report['path']).name
            assert path.stat().st_size == report['bytes']
            assert hashlib.sha256(path.read_bytes()).hexdigest() == report['sha256']
            raw = load(path)
            count_reports += 1
            count_samples += len(raw['samples'])
            for sample in raw['samples']:
                assert sample['source_sha256'] == raw['source']['sha256']
                assert sample['output'] == raw['source']
                assert sample['verification']['semantic_check'] is True
                assert sample['verification']['reopened'] is True
    assert (count_reports, count_samples) == (50, 130)
    for repeat in range(2):
        for shape in ['tiny', 'medium', 'large', 'vendor', 'unicode-vendor']:
            prefix = P / 'decoded' / f'{repeat}-{shape}'
            whole = stacks(Path(str(prefix) + '-whole.stacks'))
            owner = stacks(Path(str(prefix) + '-owner.stacks'))
            assert owner == {s: c for s, c in whole.items() if owns(s)}
            total = sum(whole.values())
            observed = sum(owner.values())
            for scope in ['whole', 'owner']:
                hist = Path(str(prefix) + f'-{scope}.histogram')
                assert sum(int(line.split()[1]) for line in hist.read_text().splitlines()) == total
                log = Path(str(prefix) + f'-{scope}.log').read_text()
                assert int(re.search(r'calls to allocation functions:\s*(\d+)', log)[1]) == total
            control = load(P / 'controls' / f'{repeat}-{shape}-allocation.json')
            profile = load(P / 'controls' / f'{repeat}-{shape}-profile-allocation.json')
            counters = []
            for report in [control, profile]:
                for sample in report['samples']:
                    a = sample['allocation']
                    assert a['failed_allocation_calls'] == 0
                    counters.append({**{k: a[k] for k in ['allocation_calls', 'reallocation_calls', 'deallocation_calls', 'allocated_bytes', 'deallocated_bytes']}, 'net_live': a['live_bytes_after'] - a['live_bytes_before'], 'peak_above_entry': a['region_peak_live_bytes'] - a['live_bytes_before']})
            assert all(c == counters[0] for c in counters)
            a = counters[0]
            assert observed == a['allocation_calls']
            assert observed != a['allocation_calls'] + a['reallocation_calls']
            rows.append({'repeat': repeat, 'shape': shape, 'whole_calls': total,
                         'owner_calls': observed, 'operation_counters': a,
                         'frozen_expected': a['allocation_calls'] + a['reallocation_calls'],
                         'frozen_pass': False, 'supplementary_counter_match': True,
                         'nested_duplicate_checks': sum(c for s,c in owner.items() if 'check_for_duplicates' in s),
                         'nested_notes_inspector': sum(c for s,c in owner.items() if 'notes::codec::inspect_element' in s)})
    result = {'reports': count_reports, 'samples': count_samples, 'rows': rows,
              'frozen_qualification': 'fail', 'owner_fractions_authorized': False}
    encoded = json.dumps(result, indent=2, sort_keys=True) + '\n'
    if args.write:
        (P / 'root-audit.json').write_text(encoded)
    if args.check:
        assert (P / 'root-audit.json').read_text() == encoded
    print('0793 independent raw audit PASS; frozen qualification FAIL in all ten traces')

if __name__ == '__main__':
    main()
