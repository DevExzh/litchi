#!/usr/bin/env python3
"""Render complete matched-source measurements and diagnostic profile scope."""
import argparse,json
from pathlib import Path
ROOT=Path(__file__).resolve().parent
def render():
    v=json.loads((ROOT/'summary.json').read_text());profiles=json.loads((ROOT/'profile-summary.json').read_text())
    lines=['# OPC Part-addition: observed and plain sources','',
      'One build, 24 reports/720 retained samples, CPU 2 and one worker, 30 samples',
      'and three warmups. Order: observed R1, plain R1, plain R2, observed R2.',
      'Both execute the same timed publication body and use a hashing discard sink.',
      'The comparison isolates the source observer; it is not a production speedup.','',
      '| Source | Mode | Shape | Repeat | p50 ms | p95 ms | p99 ms | p50 bootstrap 95% ms | Process peak RSS MiB |',
      '| --- | --- | --- | --- | ---: | ---: | ---: | --- | ---: |']
    for row in v['rows']:
        e=row['elapsed_ns'];ci=e['bootstrap_95']['p50']
        lines.append(f"| {row['source_mode']} | {row['mode']} | {row['shape']} | {row['repeat']} | {e['p50']/1e6:.6f} | {e['p95']/1e6:.6f} | {e['p99']/1e6:.6f} | {ci['lower']/1e6:.6f}–{ci['upper']/1e6:.6f} | {row['process_memory']['gnu_time_peak_rss_bytes']/2**20:.3f} |")
    lines+=['','p50 uses the midpoint; p95/p99 use nearest rank. Bootstrap intervals use',
      '2,000 deterministic resamples within each invocation, not across machines',
      'or days. Allocator elapsed time includes allocator instrumentation.','',
      '| Mode | Shape | Repeat | Plain p50 relative to observed |',
      '| --- | --- | --- | ---: |']
    for row in v['comparisons']:
        if row['metric']=='elapsed_ns.p50':lines.append(f"| {row['mode']} | {row['shape']} | {row['repeat']} | {row['relative_percent']:+.3f}% |")
    lines+=['','All paired tail/RSS/allocation differences remain in summary.json.','',
      '| Source | Shape | Repeat | Allocation calls | Requested bytes | Peak above entry | Entry live bytes | Endpoint live delta |',
      '| --- | --- | --- | ---: | ---: | ---: | ---: | ---: |']
    for row in v['rows']:
        if row['mode']!='allocator':continue
        a=row['allocation']['statistics'];d=row['allocation']['derived_statistics']
        lines.append(f"| {row['source_mode']} | {row['shape']} | {row['repeat']} | {a['allocation_calls']['p50']:,} | {a['allocated_bytes']['p50']:,} | {d['region_peak_above_entry']['p50']:,} | {a['live_bytes_before']['p50']:,} | {d['live_delta']['p50']:,} |")
    flags=[row for row in v['repeat_review'] if row['flagged']]
    lines+=['',f"Repeat review: **{len(flags)} flags** above the frozen absolute 5% trigger."]
    for row in flags:lines.append(f"- {row['source_mode']}/{row['mode']}/{row['shape']} {row['metric']}: {row['relative_percent']:+.3f}% R2 versus R1.")
    lines+=['','Every repeat check, including unflagged checks, is retained. No timed output',
      'archive is retained. Entry live bytes include preexisting input, fixture/oracle',
      'buffers and different source wrappers. Above-entry peaks and endpoint deltas',
      'must be interpreted separately from whole-process RSS.','',
      '## Stack evidence','',
      'Whole-process profiles include fixture construction, gates, warmups, timed',
      'samples and reporting. The run-frame subset still includes source setup,',
      'hash finalization and endpoint probes outside elapsed time. It is useful',
      'for choosing follow-up work, not exact attribution to the timed interval.','']
    for p in profiles['profiles']:
        lines.append(f"{p['source_mode']}: {p['sample_blocks']:,} stack blocks; {p['run_frame_blocks']:,} contain the run frame ({p['run_frame_percent']:.3f}% of sampled period).")
        lines+=['','| Run-frame leaf symbol | Weighted self % within run subset |','| --- | ---: |']
        for row in p['run_frame_self'][:8]:lines.append(f"| `{row['symbol']}` | {row['percent']:.3f}% |")
        lines+=['','Inclusive run-frame rows overlap and must not be added.','', '| Inclusive symbol | Weighted % within run subset |','| --- | ---: |']
        for row in p['run_frame_inclusive'][:10]:lines.append(f"| `{row['symbol']}` | {row['percent']:.3f}% |")
        lines.append('')
        if p['unparsed_lines']:lines.append(f"Unparsed/warning lines: {len(p['unparsed_lines'])}; retained in profile-summary.json.")
    lines+=['','Raw perf stat events, reports and stacks are retained for both modes. A zero',
      'L1 event on this guest is not proof of zero cache misses. No native/cold/range,',
      'scaling, bounded-total-memory or production CPU improvement claim follows.',
      'Plain source read/codec counters are unavailable, while its actual sink,',
      'process and allocator observations remain represented.','']
    return '\n'.join(lines)
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--check',action='store_true');a=p.parse_args();value=render();path=ROOT/'measurements.md'
    if a.check:assert path.read_text()==value
    else:
        with path.open('x') as stream:stream.write(value)
    print('VALID')
