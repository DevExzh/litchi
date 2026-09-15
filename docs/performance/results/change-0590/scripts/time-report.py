#!/usr/bin/env python3
"""Summarize paired timing legs for one selector."""
import json, sys, statistics, pathlib

S = pathlib.Path('/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0590')

def load(case, leg):
    data = json.loads((S / f'time-{case}-{leg}.json').read_text())
    for row in data['results']:
        if row['case'] == case:
            return row['elapsed_ns']['samples']
    raise SystemExit(f'case {case} missing from {leg}')

def pct(values, q):
    values = sorted(values)
    if q == 50:
        return statistics.median(values)
    idx = min(len(values) - 1, int(round((q / 100.0) * len(values) + 0.5)) - 1)
    return values[idx]

def stats(values):
    return dict(n=len(values), p50=pct(values, 50), mean=statistics.fmean(values),
                p95=pct(values, 95), p99=pct(values, 99), min=min(values), max=max(values))

def main(case):
    legs = {leg: load(case, leg) for leg in ('A1', 'B1', 'B2', 'A2')}
    before = legs['A1'] + legs['A2']
    after = legs['B1'] + legs['B2']
    print(f'case\t{case}')
    for leg in ('A1', 'B1', 'B2', 'A2'):
        s = stats(legs[leg])
        print(f'leg\t{leg}\tn={s["n"]}\tp50={s["p50"]/1e6:.3f}ms\tmean={s["mean"]/1e6:.3f}ms'
              f'\tp95={s["p95"]/1e6:.3f}ms\tp99={s["p99"]/1e6:.3f}ms')
    for label, values in (('before(A1+A2)', before), ('after(B1+B2)', after)):
        s = stats(values)
        print(f'pooled\t{label}\tn={s["n"]}\tp50={s["p50"]/1e6:.3f}ms\tmean={s["mean"]/1e6:.3f}ms'
              f'\tp95={s["p95"]/1e6:.3f}ms\tp99={s["p99"]/1e6:.3f}ms')
    b50, a50 = pct(before, 50), pct(after, 50)
    print(f'delta\tp50\tafter_vs_before={100*(a50-b50)/b50:+.2f}%\tbefore_vs_after={100*(b50-a50)/a50:+.2f}%')
    for q in (95, 99):
        bq, aq = pct(before, q), pct(after, q)
        print(f'delta\tp{q}\tafter_vs_before={100*(aq-bq)/bq:+.2f}%\tbefore_vs_after={100*(bq-aq)/aq:+.2f}%')
    bm, am = statistics.fmean(before), statistics.fmean(after)
    print(f'delta\tmean\tafter_vs_before={100*(am-bm)/bm:+.2f}%\tbefore_vs_after={100*(bm-am)/am:+.2f}%')
    a1, a2 = pct(legs['A1'], 50), pct(legs['A2'], 50)
    b1, b2 = pct(legs['B1'], 50), pct(legs['B2'], 50)
    print(f'floor\tA/A_p50\tA2_vs_A1={100*(a2-a1)/a1:+.2f}%')
    print(f'floor\tB/B_p50\tB2_vs_B1={100*(b2-b1)/b1:+.2f}%')
    for q in (95, 99):
        qa1, qa2 = pct(legs['A1'], q), pct(legs['A2'], q)
        print(f'floor\tA/A_p{q}\tA2_vs_A1={100*(qa2-qa1)/qa1:+.2f}%')

if __name__ == '__main__':
    main(sys.argv[1])
