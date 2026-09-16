"""Paired timing summary for change 0598 (A1 B1 B2 A2)."""
import json, sys, statistics

S = sys.argv[1] if len(sys.argv) > 1 else '/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0598'

def load(tag):
    d = json.load(open(f'{S}/timing-{tag}.json'))
    out = {}
    for r in d['results']:
        cc = r['source'].get('pptx_cross_copy') or r['source'].get('pptx_cross_copy_lifecycle') or {}
        out[r['case']] = {
            'samples': r['elapsed_ns']['samples'],
            'phases': {k: v for k, v in cc.items() if k.endswith('_ns')},
            'digest': cc.get('expected_output_sha256'),
            'outputs': sorted(set(cc.get('output_sha256', []))),
        }
    return out, d['binary_identity']['binary_sha256']

def q(xs, p):
    xs = sorted(xs)
    if not xs:
        return float('nan')
    k = min(len(xs) - 1, max(0, int(round((p / 100.0) * (len(xs) - 1)))))
    return xs[k]

def stats(xs):
    return dict(p50=q(xs, 50), mean=statistics.fmean(xs), p95=q(xs, 95), p99=q(xs, 99), min=min(xs), n=len(xs))

def pct(a, b):
    return 100.0 * (a - b) / b

legs = {}
shas = {}
for tag in ('A1', 'B1', 'B2', 'A2'):
    legs[tag], shas[tag] = load(tag)

print('binary sha256 per leg:')
for tag in ('A1', 'B1', 'B2', 'A2'):
    print(f'  {tag}: {shas[tag]}')

for case in legs['A1']:
    print(f'\n################ {case} ################')
    per = {tag: stats(legs[tag][case]['samples']) for tag in legs}
    before = legs['A1'][case]['samples'] + legs['A2'][case]['samples']
    after = legs['B1'][case]['samples'] + legs['B2'][case]['samples']
    bs, as_ = stats(before), stats(after)
    print(f'{"leg":>6s} {"n":>4s} {"p50 ms":>12s} {"mean ms":>12s} {"p95 ms":>12s} {"p99 ms":>12s}')
    for tag in ('A1', 'B1', 'B2', 'A2'):
        s = per[tag]
        print(f'{tag:>6s} {s["n"]:>4d} {s["p50"]/1e6:>12.4f} {s["mean"]/1e6:>12.4f} {s["p95"]/1e6:>12.4f} {s["p99"]/1e6:>12.4f}')
    print(f'{"BEFORE":>6s} {bs["n"]:>4d} {bs["p50"]/1e6:>12.4f} {bs["mean"]/1e6:>12.4f} {bs["p95"]/1e6:>12.4f} {bs["p99"]/1e6:>12.4f}')
    print(f'{"AFTER":>6s} {as_["n"]:>4d} {as_["p50"]/1e6:>12.4f} {as_["mean"]/1e6:>12.4f} {as_["p95"]/1e6:>12.4f} {as_["p99"]/1e6:>12.4f}')
    for k in ('p50', 'mean', 'p95', 'p99'):
        print(f'  after vs before {k}: {pct(as_[k], bs[k]):+.2f}%   before vs after: {pct(bs[k], as_[k]):+.2f}%')
    aa = {k: pct(per['A2'][k], per['A1'][k]) for k in ('p50', 'mean', 'p95', 'p99')}
    bb = {k: pct(per['B2'][k], per['B1'][k]) for k in ('p50', 'mean', 'p95', 'p99')}
    print(f'  A/A floor (A2 vs A1): ' + ', '.join(f'{k} {v:+.2f}%' for k, v in aa.items()))
    print(f'  B/B floor (B2 vs B1): ' + ', '.join(f'{k} {v:+.2f}%' for k, v in bb.items()))
    outs = set()
    for tag in legs:
        outs.update(legs[tag][case]['outputs'])
    print(f'  distinct output_sha256 across all four legs: {sorted(outs)}')
    phases = sorted({p for tag in legs for p in legs[tag][case]['phases']})
    if phases:
        print(f'  {"phase":>18s} {"before p50 us":>16s} {"after p50 us":>16s} {"delta":>10s}')
        for p in phases:
            b = legs['A1'][case]['phases'].get(p, []) + legs['A2'][case]['phases'].get(p, [])
            a = legs['B1'][case]['phases'].get(p, []) + legs['B2'][case]['phases'].get(p, [])
            if not b or not a:
                continue
            bp, ap = q(b, 50), q(a, 50)
            d = f'{pct(ap, bp):+9.2f}%' if bp else '        --'
            print(f'  {p:>18s} {bp/1e3:>16.3f} {ap/1e3:>16.3f} {d}')
