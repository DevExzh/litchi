#!/usr/bin/env python3
"""Change 0750: total the differential campaign's per-seed counters."""
import glob, re, sys
keys = ['cases', 'original', 'replacement', 'identical', 'window', 'complete', 'two_sites', 'two_site_windows']
totals = dict.fromkeys(keys, 0)
seeds = 0
for path in sorted(glob.glob(sys.argv[1] + '/seed*.log')):
    text = open(path).read()
    m = re.search(r'cases (\d+): refused on the original (\d+), refused on the replacement (\d+), '
                  r'identical (\d+), window (\d+), complete (\d+); two separated edits (\d+) \(window (\d+)\)', text)
    ok = 'test result: ok' in text
    print(path.rsplit('/', 1)[-1], 'ok' if ok else 'FAILED', m.group(0) if m else '')
    if m and ok:
        seeds += 1
        for key, value in zip(keys, m.groups()):
            totals[key] += int(value)
print('seeds passed', seeds, totals)
