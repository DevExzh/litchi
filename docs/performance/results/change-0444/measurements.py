#!/usr/bin/env python3
"""Render all retained rows and reviewed limitations from the baseline summary."""
import argparse,json
from pathlib import Path
ROOT=Path(__file__).resolve().parent
def render():
    v=json.loads((ROOT/'summary.json').read_text())
    lines=['# OPC Part-addition observed baseline','',
      'One current revision; 12 reports, 360 retained samples, CPU 2, one worker,',
      '30 samples and three warmups per report. Normal mode includes the ReadAt',
      'observer and hashing sink. Allocator elapsed values include instrumentation.',
      'These are not plain-source production latency estimates or speedup claims.','',
      '| Mode | Shape | Repeat | p50 ms | p95 ms | p99 ms | p50 bootstrap 95% ms | Whole-process peak RSS MiB |',
      '| --- | --- | --- | ---: | ---: | ---: | --- | ---: |']
    for r in v['rows']:
        e=r['elapsed_ns'];ci=e['bootstrap_95']['p50']
        lines.append(f"| {r['mode']} | {r['shape']} | {r['repeat']} | {e['p50']/1e6:.6f} | {e['p95']/1e6:.6f} | {e['p99']/1e6:.6f} | {ci['lower']/1e6:.6f}–{ci['upper']/1e6:.6f} | {r['process_memory']['gnu_time_peak_rss_bytes']/2**20:.3f} |")
    lines+=['','p50 uses the midpoint of the central observations; p95/p99 use nearest rank.',
      'Intervals use 2,000 deterministic bootstrap resamples within each invocation;',
      'they do not measure machine-to-machine or day-to-day uncertainty. Raw samples,',
      'all quantiles/intervals and every repeat check are retained in summary.json.','',
      '| Shape | Allocation calls | Requested bytes | Region peak above entry | Endpoint live delta | Source calls | Returned source bytes | Output bytes |',
      '| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |']
    for r in v['rows']:
        if r['mode']!='allocator' or r['repeat']!='R1':continue
        a=r['allocation']['statistics'];d=r['allocation']['derived_statistics']
        report=json.loads((ROOT/r['report']).read_text())['results'][0];source=report['source']
        lines.append(f"| {r['shape']} | {a['allocation_calls']['p50']:,} | {a['allocated_bytes']['p50']:,} | {d['region_peak_above_entry']['p50']:,} | {d['live_delta']['p50']:,} | {source['read_calls'][0]:,} | {source['read_bytes'][0]:,} | {report['sink']['accepted_bytes']:,} |")
    flags=[r for r in v['repeat_review'] if r['flagged']]
    lines+=['',f"Repeat review: **{len(flags)} flags** above the frozen absolute 5% trigger.",
      'Allocation values and source/sink counters match both repeats. No timed output',
      'archive is retained; the endpoint live delta is zero because publication',
      'consumes and drops its package. Preexisting input/oracle fixtures stay live.',
      'Metadata/temporary peaks grow with Part count, so no bounded-total-memory',
      'claim follows. Source calls include repeated reads and observer bookkeeping;',
      'codec byte-flow, remote requests, lock wait and scaling are not measured.','',
      '## Profile review','',
      'The large normal whole-process profile reports **55.24% self time in the',
      'instrumented source reader**, followed by SHA-256 at 9.79%. The reader scans',
      'all ordinary member ranges for each read. This observer work can grow with',
      'both read count and member count; the curve cannot establish production',
      'topology complexity. A matched plain-source lifecycle is the next measurement',
      'priority before a production optimization. No Amdahl speedup is inferred',
      'from the whole-process fraction, which includes fixture construction, gates,',
      'warmups, hashing and report work outside the timed interval.','',
      'Perf stat returned 11,538,555,261 cycles, 49,721,657,346 instructions,',
      '4,964,600,258 branches and 8,676,980 branch misses. Its L1-miss event returned',
      'zero on this guest; that is not evidence of zero hardware cache misses.',
      'Perf report/script retained addr2line warnings. Symbol-level self rows are',
      'usable, but precise source-line/inlined attribution remains limited. No',
      'samples were reported lost. The raw profile, warnings and commands remain',
      'in the bundle; profiles are diagnostic, not paired CPU improvement evidence.','']
    return '\n'.join(lines)
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--check',action='store_true');a=p.parse_args();text=render();target=ROOT/'measurements.md'
    if a.check:assert target.read_text()==text
    else:
        with target.open('x') as stream:stream.write(text)
    print('VALID')
