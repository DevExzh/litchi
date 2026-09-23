#!/usr/bin/env python3
"""Markdown tables for record 0757 from the retained JSON summaries.

usage: tables.py PACKET_DIR  (the change-0757 packet directory)
"""
import json, os, sys

D = sys.argv[1]


def load(path):
    with open(os.path.join(D, path)) as handle:
        return json.load(handle)


def latency():
    rows = {}
    for window in ('w1', 'w2'):
        for name in ('summary-xls.json', 'summary-docppt.json'):
            for r in load(f'latency/window-{window[1]}/{name}'):
                rows.setdefault((r['case'], r['shape']), {})[window] = r
    print('| selector | shape | samples | before p50 ms | after p50 ms | window 1 ratio [95% CI] | window 2 ratio [95% CI] | same output |')
    print('| --- | --- | ---: | ---: | ---: | ---: | ---: | --- |')
    for (case, shape), w in rows.items():
        a, b = w['w1'], w['w2']
        same = all(x['archive_sha256_before'] == x['archive_sha256_after'] for x in (a, b))
        fmt = lambda v: f'{v / 1e6:.4f}' if v < 1e7 else f'{v / 1e6:.3f}'
        print(f"| `{case}` | {shape} | {a['warmup']}+{a['samples']} | {fmt(a['before_median_p50'])} | {fmt(a['after_median_p50'])} | "
              f"**{a['median_paired_p50_ratio']:.3f}** [{a['bootstrap_95ci_p50_ratio'][0]:.3f}, {a['bootstrap_95ci_p50_ratio'][1]:.3f}] | "
              f"{b['median_paired_p50_ratio']:.3f} [{b['bootstrap_95ci_p50_ratio'][0]:.3f}, {b['bootstrap_95ci_p50_ratio'][1]:.3f}] | {same} |")


def counters():
    rows = []
    for name in sorted(os.listdir(os.path.join(D, 'counters'))):
        if name.startswith('summary'):
            rows += load(f'counters/{name}')
    print('| selector | shape | user instructions | user cycles | kernel instructions | page faults |')
    print('| --- | --- | --- | --- | --- | --- |')
    for r in rows:
        a, b = r['A'], r['B']
        def cell(key, digits=0):
            change = f' ({b[key] / a[key] - 1:+.1%})' if abs(a[key]) > 1 else ''
            return f"{a[key]:,.{digits}f} → {b[key]:,.{digits}f}{change}"
        print(f"| `{r['case']}` | {r['shape']} | {cell('instructions:u')} | {cell('cycles:u')} | {cell('instructions:k')} | {cell('page-faults')} |")


def probe(path, title):
    print(f'\n{title}\n')
    print('| probe case | writes per process | processes | before p50 ms | after p50 ms | paired ratio [95% CI] | instructions per write (callgrind) | distinct outputs before / after |')
    print('| --- | ---: | ---: | ---: | ---: | ---: | ---: | --- |')
    for r in load(path):
        i = r['instructions_per_write']
        print(f"| `{r['case']}` | {r['iterations']} | {len(r['processes'])} | {r['before_median_p50_ns'] / 1e6:.4f} | {r['after_median_p50_ns'] / 1e6:.4f} | "
              f"**{r['median_paired_p50_ratio']:.3f}** [{r['bootstrap_95ci_p50_ratio'][0]:.3f}, {r['bootstrap_95ci_p50_ratio'][1]:.3f}] | "
              f"{i['A']:,.0f} → {i['B']:,.0f} ({r['instructions_ratio']:.3f}) | {len(r['outputs_before'])} / {len(r['outputs_after'])} |")


def allocations():
    print('\n| probe case | allocations | reallocations | allocated bytes | peak live bytes |')
    print('| --- | --- | --- | --- | --- |')
    for r in load('probe/summary.json'):
        a, b = r['allocations_before'], r['allocations_after']
        def cell(key):
            change = f' ({b[key] / a[key] - 1:+.1%})' if a[key] else ''
            return f'{a[key]:,} → {b[key]:,}{change}'
        print(f"| `{r['case']}` | {cell('allocations_per_write')} | {cell('reallocations_per_write')} | {cell('allocated_bytes_per_write')} | {cell('peak_live_bytes')} |")


def tunables():
    print('\n| selector / shape | malloc thresholds | leg order A B B A: timed p50 ms | page faults per iteration | user instructions per iteration |')
    print('| --- | --- | --- | --- | --- |')
    for name in sorted(os.listdir(os.path.join(D, 'tunables'))):
        if not name.startswith('summary'):
            continue
        rows = load(f'tunables/{name}')
        case = name[len('summary-'):-len('.json')]
        for setting in ('default', 'pinned'):
            chosen = [r for r in rows if r['setting'] == setting]
            p50 = ' / '.join(f"{r['leg']} {r['p50_ms']:.3f}" for r in chosen)
            faults = ' / '.join(f"{r['page-faults']:.0f}" for r in chosen)
            instr = ' / '.join(f"{r['instructions:u'] / 1e6:.2f}M" for r in chosen)
            print(f'| `{case}` | {setting} | {p50} | {faults} | {instr} |')


if __name__ == '__main__':
    print('## latency'); latency()
    print('\n## counters'); counters()
    probe('probe/summary.json', '## probe (8 processes, default malloc)')
    probe('probe-multi-default/summary-default-4r.json', '## probe multi-string (16 processes, default malloc)')
    probe('probe-multi-pinned/summary-pinned-4r.json', '## probe multi-string (16 processes, pinned malloc thresholds)')
    print('\n## allocations'); allocations()
    print('\n## tunables'); tunables()
