#!/usr/bin/env python3
"""Per-lifecycle isolation pairs for change 0645: totals, call counts, and
inclusive Ir of the hashing subtree, across the before and scratch legs."""
import pathlib, sys
S = pathlib.Path('/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0645/out')
CASES = ['pptx_eager_batch_edit_save', 'pptx_eager_multi_slide_batch_edit_save',
         'pptx_slide_move_boundary_save', 'pptx_slide_remove_boundary_save']
LEGS = ['before', 'after', 'afterb']
FP = ('package_fingerprint', 'package_fingerprint_with_memo')

def calls(tag):
    d = {}
    for line in (S / f'calls-{tag}.txt').read_text().splitlines():
        p = line.split('\t')
        if p[0] == 'total_Ir':
            d['total'] = int(p[1])
        elif p[0] == 'calls' and len(p) > 2 and 'no callee' not in p[2]:
            k = p[2].split('::')[-1]
            d[k] = d.get(k, 0) + int(p[1])
    return d

def incl(tag, needle):
    best = 0
    for line in (S / f'incl-{tag}.txt').read_text().splitlines():
        t = line.strip()
        if needle in t and t and t[0].isdigit():
            best = max(best, int(t.split()[0].replace(',', '')))
    return best

def fingerprint_incl(tag):
    """Top-level fingerprint subtree: the v2 leg nests the thin wrapper inside
    the memoizing entry point, so the memoizing total already covers it."""
    memo = incl(tag, 'package_fingerprint_with_memo')
    return memo if memo else incl(tag, 'package_fingerprint')

for case in CASES:
    print(f'## {case}')
    base = None
    for leg in LEGS:
        try:
            c1, c3 = calls(f'{leg}-{case}-s1'), calls(f'{leg}-{case}-s3')
        except FileNotFoundError:
            print(f'  {leg}: (missing)'); continue
        life = (c3['total'] - c1['total']) // 2
        fp1, fp3 = fingerprint_incl(f'{leg}-{case}-s1'), fingerprint_incl(f'{leg}-{case}-s3')
        fplife = (fp3 - fp1) // 2
        sha1, sha3 = incl(f'{leg}-{case}-s1', 'compress256'), incl(f'{leg}-{case}-s3', 'compress256')
        shalife = (sha3 - sha1) // 2
        fpcalls = sum(((c3.get(k, 0) - c1.get(k, 0)) / 2) for k in FP)
        capcalls = (c3.get('capture_internal', 0) - c1.get('capture_internal', 0)) / 2
        eqcalls = (c3.get('packages_equal', 0) - c1.get('packages_equal', 0)) / 2
        if base is None:
            base = (life, fplife, shalife)
        d = lambda x, i: f'{100*(x-base[i])/base[i]:+7.2f}%' if base[i] else '      -'
        print(f'  {leg:7s} Ir/life {life:15,d} {d(life,0)}   fingerprint {fplife:13,d} {d(fplife,1)}'
              f'   sha2 {shalife:13,d} {d(shalife,2)}   fp_calls {fpcalls:4.1f} captures {capcalls:4.1f} eq {eqcalls:4.1f}')
    print()
