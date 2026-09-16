#!/usr/bin/env python3
"""Inclusive Ir and call counts per (caller, callee) pair from a callgrind.out.

Callgrind emits, after each `cfn=` / `calls=` pair, one cost line whose first
number is the inclusive cost of those calls. Summing per caller separates the
call sites of one callee.
"""
import re, sys, collections

def main(path, needle):
    names = {}
    caller = None
    pending = None
    pending_calls = 0
    per = collections.Counter()
    cnt = collections.Counter()
    for line in open(path, errors='replace'):
        line = line.rstrip('\n')
        m = re.match(r'^fn=\((\d+)\)(?:\s+(.*))?$', line)
        if m:
            if m.group(2):
                names[m.group(1)] = m.group(2)
            caller = names.get(m.group(1), '?')
            pending = None
            continue
        m = re.match(r'^cfn=\((\d+)\)(?:\s+(.*))?$', line)
        if m:
            if m.group(2):
                names[m.group(1)] = m.group(2)
            pending = names.get(m.group(1), '?')
            continue
        m = re.match(r'^calls=(\d+)', line)
        if m and pending is not None:
            pending_calls = int(m.group(1))
            continue
        if pending is not None and pending_calls and re.match(r'^[\d+*-]', line):
            parts = line.split()
            if len(parts) >= 2 and parts[1].lstrip('+-').isdigit():
                if needle in pending:
                    per[(caller, pending)] += int(parts[1])
                    cnt[(caller, pending)] += pending_calls
            pending_calls = 0
            pending = None
            continue
    for key, ir in sorted(per.items(), key=lambda kv: -kv[1]):
        caller, callee = key
        n = cnt[key]
        print(f'{ir:16,d}\t{n:6d}\t{ir//max(n,1):14,d}\t{caller[:70]}\t->\t{callee[:60]}')

if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
