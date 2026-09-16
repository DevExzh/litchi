#!/usr/bin/env python3
"""Sum callgrind call counts and inclusive Ir per callee name from a callgrind.out."""
import re, sys, collections

def main(path, names):
    calls = collections.Counter()
    ir = collections.Counter()
    names_by_id = {}
    pending = None
    total = 0
    with open(path, 'r', errors='replace') as handle:
        for line in handle:
            line = line.rstrip('\n')
            m = re.match(r'^cfn=\((\d+)\)(?:\s+(.*))?$', line)
            if m:
                idx, nm = m.group(1), m.group(2)
                if nm:
                    names_by_id[idx] = nm
                pending = names_by_id.get(idx, '?')
                continue
            m = re.match(r'^fn=\((\d+)\)(?:\s+(.*))?$', line)
            if m:
                idx, nm = m.group(1), m.group(2)
                if nm:
                    names_by_id[idx] = nm
                pending = None
                continue
            m = re.match(r'^calls=(\d+)', line)
            if m and pending is not None:
                calls[pending] += int(m.group(1))
                continue
            m = re.match(r'^summary:\s+(\d+)', line)
            if m:
                total = int(m.group(1))
    print(f"total_Ir\t{total}")
    for needle in names:
        hits = [(k, v) for k, v in calls.items() if needle in k]
        hits.sort(key=lambda kv: -kv[1])
        for k, v in hits[:6]:
            print(f"calls\t{v}\t{k}")
        if not hits:
            print(f"calls\t0\t<no callee matching {needle}>")

if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2:])
