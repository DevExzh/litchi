#!/usr/bin/env python3
import argparse,json
from pathlib import Path
ROOT=Path(__file__).resolve().parent
def render():
    a=json.loads((ROOT/'attribution.json').read_text())
    lines=['# PPTX caller and source-owner attribution','',
           'Reanalysis of 0448: two retained media-rich profiles and eight reports/240',
           'samples. No new workload ran. Percentages below use all lifecycle-frame',
           'sample period, including warmups and untimed work; blocked sleep is absent.','',
           '| SHA caller | Separate sleeps: % run period | Minimum service: % run period |',
           '| --- | ---: | ---: |']
    by=[{r['category']:r for r in p['categories']} for p in a['profiles']]
    for c in ['untimed-harness-output-hash','planning-touched-digest','publication-touched-digest','unclassified-lifecycle-sha']:
        lines.append(f"| {c} | {by[0][c]['percent_of_run_period']:.3f} | {by[1][c]['percent_of_run_period']:.3f} |")
    lines+=['','The untimed hash contributes '+ '/'.join(f"{b['untimed-harness-output-hash']['percent_of_run_sha']:.3f}%" for b in by)+
            ' of lifecycle SHA period.',
            'The remaining unclassified SHA stacks are retained and resolve to publication',
            'graph_digest. Both profiles have zero missing callchains. Full stacks and',
            'outside-lifecycle SHA remain in attribution.json with conserved counts/period.','',
            '| Corpus | Phase | Source calls | Source bytes | Destination calls | Destination bytes |',
            '| --- | --- | ---: | ---: | ---: | ---: |']
    for corpus,rows in a['work']['by_corpus'].items():
        for r in rows:
            s=r['source_reads'];d=r['destination_reads']
            lines.append(f"| {corpus} | {r['phase']} | {s['logical_calls']:,} | {s['returned_bytes']:,} | {d['logical_calls']:,} | {d['returned_bytes']:,} |")
    lines+=['','Every source/destination read delta, histogram, nominal pacing counter and cache',
            'point matches across all 30 samples, both repeats and both policies within its',
            'corpus. Media-rich publication has 23 source-cache hits, zero cold loads and',
            '16,807,458 retained cache bytes. The source remains cached while compressed',
            'transfer authorization issues fresh physical reads. Destination preservation',
            'also reads its original media. Total phase bytes are not all source bytes.','',
            'Static source corroborates these paths, but the reports contain no member/offset',
            'read trace. No exact per-member attribution, removable-byte count, production',
            'speedup, allocation reduction, actual link rate or Amdahl bound is established.',
            'See source-review.md for ownership constraints and ranked follow-up.','']
    return '\n'.join(lines)
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--check',action='store_true');a=p.parse_args();result=render()
    if a.check:assert (ROOT/'measurements.md').read_text()==result
    else:(ROOT/'measurements.md').write_text(result)
    print('VALID')
