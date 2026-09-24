#!/usr/bin/env python3
"""For each path: lines added by HEAD (ours, stage 2) or incoming (theirs, stage 3)
relative to base (stage 1) that are missing from the working-tree file.
Uses multiset diff; ignores blank/brace-only lines."""
import sys, subprocess
from collections import Counter
def show(rev, path):
    r = subprocess.run(['git', 'show', f'{rev}:{path}'], capture_output=True, text=True)
    return r.stdout.split('\n') if r.returncode == 0 else []
trivial = lambda l: l.strip() in ('', '}', '{', '},', ')', '});', '};', ']', '),', ');', '})', '}))')
for path in sys.argv[1:]:
    base = Counter(show('f0ab67b55d', path)); ours = Counter(show('e6cca92db2', path)); theirs = Counter(show('a67a38abf2', path))
    cur = Counter(open(path, encoding='utf-8').read().split('\n'))
    for name, side in (('HEAD', ours), ('INC', theirs)):
        added = side - base
        missing = [l for l in (added - cur).elements() if not trivial(l)]
        if missing:
            print(f'== {path}: {name} lines missing: {len(missing)}')
            for l in missing[:40]:
                print('   ', l[:160])
