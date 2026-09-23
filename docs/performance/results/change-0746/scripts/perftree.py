#!/usr/bin/env python3
"""Top-down inclusive call tree from `perf script` output (frame-pointer
callchains with inlined frames), merged by function name.

usage: perftree.py PERF.DATA [--root NAME] [--min PCT] [--depth N]
Percentages are of all samples in the file; --root restricts the tree to
the subtree(s) under the first frame (from the root side) whose name
equals NAME.
"""
import subprocess, sys, re, collections, argparse

ap = argparse.ArgumentParser()
ap.add_argument('data')
ap.add_argument('--root', default=None)
ap.add_argument('--min', type=float, default=1.0)
ap.add_argument('--depth', type=int, default=12)
ap.add_argument('--collapse', default='branch,from_residual,map,map_err,and_then,unwrap_or_else,call_once,{closure#0},ok_or_else,try_fold,fold,next,for_each,__rust_begin_short_backtrace')
args = ap.parse_args()
collapse = set(args.collapse.split(','))

out = subprocess.run(['perf', 'script', '-i', args.data, '-F', 'period,ip,sym', '--no-demangle'],
                     capture_output=True, text=True).stdout
# demangle via rustfilt-less approach: use perf's own demangling instead
out = subprocess.run(['perf', 'script', '-i', args.data, '-F', 'period,ip,sym'],
                     capture_output=True, text=True).stdout

samples = []
cur = None
for line in out.split('\n'):
    if not line.strip():
        if cur is not None:
            samples.append(cur)
        cur = None
        continue
    if cur is None:
        # header line: period
        try:
            period = int(line.strip().split()[0])
        except Exception:
            period = 1
        cur = [period, []]
        continue
    m = re.match(r'\s*[0-9a-f]+\s+(.*)$', line)
    if not m:
        continue
    sym = m.group(1)
    sym = re.sub(r'\s*\(inlined\)\s*$', '', sym)
    sym = re.sub(r'\+0x[0-9a-f]+$', '', sym)
    # strip generic args for readability
    depth = 0
    s2 = ''
    for ch in sym:
        if ch == '<':
            depth += 1
            if depth == 1:
                s2 += '<..>'
            continue
        if ch == '>':
            depth -= 1
            continue
        if depth == 0:
            s2 += ch
    s2 = re.sub(r'::h[0-9a-f]{16}$', '', s2)
    cur[1].append(s2)
if cur is not None:
    samples.append(cur)

total = sum(p for p, _ in samples)
class Node:
    __slots__ = ('w', 'kids')
    def __init__(self):
        self.w = 0
        self.kids = collections.OrderedDict()
tree = Node()
for period, frames in samples:
    chain = list(reversed(frames))  # root first
    chain = [f for f in chain if f.split('::')[-1].replace('<..>','') not in collapse and f.replace('<..>','') not in collapse]
    if args.root:
        try:
            i = chain.index(args.root)
        except ValueError:
            continue
        chain = chain[i:]
    node = tree
    node.w += period
    seen = []
    for f in chain:
        node = node.kids.setdefault(f, Node())
        node.w += period

def show(node, name, depth, indent):
    pct = 100.0 * node.w / total
    if pct < args.min or depth > args.depth:
        return
    print('%s%6.2f%%  %s' % ('  ' * indent, pct, name[:150]))
    for k, v in sorted(node.kids.items(), key=lambda kv: -kv[1].w):
        show(v, k, depth + 1, indent + 1)

print('total samples weight', total, 'rooted weight %.2f%%' % (100.0 * tree.w / total))
for k, v in sorted(tree.kids.items(), key=lambda kv: -kv[1].w):
    show(v, k, 1, 0)
