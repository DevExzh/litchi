#!/usr/bin/env python3
"""Parse callgrind output: per-function self Ir, inclusive Ir (sum of inclusive
cost over incoming calls), and incoming call counts. Uses name compression."""
import re, sys, collections

def parse(path):
    fn_names = {}
    cur_fn = None
    self_ir = collections.Counter()
    incl_in = collections.Counter()   # callee -> inclusive Ir over calls
    calls_in = collections.Counter()  # callee -> call count
    edges = collections.Counter()     # (caller, callee) -> calls
    edge_ir = collections.Counter()
    pending_call = None
    def name(tok):
        m = re.match(r'\((\d+)\)(?:\s+(.*))?', tok)
        if m:
            i, n = m.group(1), m.group(2)
            if n: fn_names[i] = n
            return fn_names[i]
        return tok
    with open(path, errors='replace') as f:
        for line in f:
            line = line.rstrip('\n')
            if line.startswith('fn='):
                cur_fn = name(line[3:])
            elif line.startswith('cfn='):
                cfn = name(line[4:])
            elif line.startswith('calls='):
                n = int(line[6:].split()[0])
                pending_call = (cfn, n)
            elif line and (line[0].isdigit() or line[0] in '+-*'):
                parts = line.split()
                if len(parts) < 2: continue
                ir = int(parts[1])
                if pending_call:
                    callee, n = pending_call
                    incl_in[callee] += ir
                    calls_in[callee] += n
                    edges[(cur_fn, callee)] += n
                    edge_ir[(cur_fn, callee)] += ir
                    pending_call = None
                else:
                    self_ir[cur_fn] += ir
            elif line.startswith('summary:') or line.startswith('totals:'):
                total = int(line.split()[1])
    return self_ir, incl_in, calls_in, edges, edge_ir, total

if __name__ == '__main__':
    for path in sys.argv[1:]:
        s, inc, c, e, eir, total = parse(path)
        print(path, 'total', total)
