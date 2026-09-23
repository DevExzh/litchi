#!/usr/bin/env python3
"""Render the record's tables from the packet's summaries (change 0753).

usage: tables.py PACKET_DIR
"""
import json, os, sys

P = sys.argv[1]


def load(path):
    with open(os.path.join(P, path)) as handle:
        return json.load(handle)


def ms(ns):
    return f'{ns / 1e6:.4f}' if ns < 1e6 else f'{ns / 1e6:.3f}'


window1 = {(r['case'], r['shape']): r for f in ('summary-fresh.json', 'summary-controls.json')
           for r in load(f'latency/window-1/{f}')}
window2 = {(r['case'], r['shape']): r for r in load('latency/window-2/summary-fresh.json')}
print('| selector | shape | before p50 ms | after p50 ms | paired ratio | 95% CI | window 2 ratio | 95% CI |')
print('| --- | --- | ---: | ---: | ---: | --- | ---: | --- |')
for key, r in window1.items():
    w2 = window2.get(key)
    lo, hi = r['bootstrap_95ci_p50_ratio']
    second = f"{w2['median_paired_p50_ratio']:.3f} | [{w2['bootstrap_95ci_p50_ratio'][0]:.3f}, {w2['bootstrap_95ci_p50_ratio'][1]:.3f}]" if w2 else '— | —'
    assert r['archive_sha256_before'] == r['archive_sha256_after']
    print(f"| `{key[0]}` | {key[1]} | {ms(r['before_median_p50'])} | {ms(r['after_median_p50'])} | "
          f"**{r['median_paired_p50_ratio']:.3f}** | [{lo:.3f}, {hi:.3f}] | {second} |")
print()
counters = {}
for name in sorted(os.listdir(os.path.join(P, 'counters'))):
    if name.startswith('summary') and name.endswith('.json'):
        for row in load(f'counters/{name}'):
            counters[(row['case'], row['shape'])] = row
print('| selector | shape | user instructions | user cycles | kernel instructions | page faults |')
print('| --- | --- | --- | --- | --- | --- |')
for key, row in counters.items():
    a, b = row['A'], row['B']
    def cell(event, digits=0):
        return f"{a[event]:,.{digits}f} → {b[event]:,.{digits}f} ({b[event] / a[event] - 1:+.1%})" if a[event] else f"{a[event]:,.0f} → {b[event]:,.0f}"
    print(f"| `{key[0]}` | {key[1]} | {cell('instructions:u')} | {cell('cycles:u')} | "
          f"{cell('instructions:k')} | {a['page-faults']:,.0f} → {b['page-faults']:,.0f} |")
print()
print('| writer / shape | instructions per write, before | after | after / before |')
print('| --- | ---: | ---: | ---: |')
for row in load('callgrind/summary.json'):
    print(f"| `{row['case']}` | {row['A']:,.0f} | {row['B']:,.0f} | **{row['ratio']:.3f}** |")
print()
a = {r['case']: r for r in load('alloc/probe-A.json')}
b = {r['case']: r for r in load('alloc/probe-B.json')}
print('| writer / shape | allocations | reallocations | allocated bytes | peak live bytes |')
print('| --- | --- | --- | --- | --- |')
for key in a:
    x, y = a[key], b[key]
    print(f"| `{key}` | {x['allocations']:,} → {y['allocations']:,} | {x['reallocations']:,} → {y['reallocations']:,} | "
          f"{x['allocated_bytes']:,} → {y['allocated_bytes']:,} ({y['allocated_bytes'] / x['allocated_bytes'] - 1:+.1%}) | "
          f"{x['peak_live_bytes']:,} → {y['peak_live_bytes']:,} ({y['peak_live_bytes'] / x['peak_live_bytes'] - 1:+.1%}) |")
