#!/usr/bin/env python3
"""Render measured simulation costs and explicit memory/timing limits."""
import argparse,json
from pathlib import Path
ROOT=Path(__file__).resolve().parent
def render():
    v=json.loads((ROOT/'summary.json').read_text());lines=['# PPTX range transfer-pacing baseline','','One build; 8 reports/240 retained samples; CPU 2, one worker, 30 samples and','three warmups. Both configurations use 64 KiB maximum reads and 200 us fixed','request delay. Paced reads additionally request transfer sleeps at 25 MiB/s.','This measures simulation cost, not a production speedup or physical network rate.','','| Corpus | Paced | Repeat | API p50 ms | p95 ms | p99 ms | p50 bootstrap 95% ms | Whole-process RSS MiB |','| --- | --- | --- | ---: | ---: | ---: | --- | ---: |']
    for row in v['rows']:
        t=row['timings']['api_sum_ns'];lo,hi=t['p50_bootstrap_95'];lines.append(f"| {row['corpus']} | {row['paced']} | {row['repeat']} | {t['p50']/1e6:.6f} | {t['p95']/1e6:.6f} | {t['p99']/1e6:.6f} | {lo/1e6:.6f}–{hi/1e6:.6f} | {row['whole_process_peak_rss_bytes']/2**20:.3f} |")
    lines+=['','All API vectors, per-phase timing and counter observations are in summary.json.','p50 uses the midpoint; p95/p99 use nearest rank. Bootstrap intervals use 2,000','deterministic within-process resamples, not machine/day uncertainty.','','| Corpus | Paced | Repeat | Phase | Logical reads | Returned bytes | Requested transfer delay ms |','| --- | --- | --- | --- | ---: | ---: | ---: |']
    for row in v['rows']:
        for phase,metrics in row['phase_read_deltas'].items():lines.append(f"| {row['corpus']} | {row['paced']} | {row['repeat']} | {phase} | {metrics['logical_calls']['p50']:,} | {metrics['returned_bytes']['p50']:,} | {metrics['transfer_delay_ns']['p50']/1e6:.6f} |")
    lines+=['','| Corpus | Repeat | Phase | Requested transfer delay / paced API median |','| --- | --- | --- | ---: |']
    for row in v['rows']:
        if row['paced']:
            for phase,clock in [('opened','open_ns'),('planned','plan_ns'),('published','publication_ns')]:
                share=100*row['phase_read_deltas'][phase]['transfer_delay_ns']['p50']/row['timings'][clock]['p50']
                lines.append(f"| {row['corpus']} | {row['repeat']} | {phase} | {share:.3f}% |")
    lines+=['','These are requested-sleep/median ratios, not Amdahl serial fractions or','measured sleep attribution. Fixed delay, OS oversleep and actual API work','remain in the denominator; no causal subtraction is performed.']
    flags=[r for r in v['repeat_review'] if r['flagged']];lines+=['',f'Absolute 5% repeat review: {len(flags)} flags.','']
    for row in flags:lines.append(f"- {row['corpus']}, paced={row['paced']}, {row['metric']}: {row['relative_percent']:+.3f}% R2 versus R1.")
    lines+=['','Every paired API difference remains visible in summary.json. Added pacing cost','is expected model behavior and is not a production regression. Counter equality','is checked across configurations, repeats and samples; output/source identities','are frozen from the pilots. No observation is silently recaptured or discarded.','','The full-retaining CountingSink and fixed fixtures contribute to process memory.','Managed phase-boundary memory is not an allocation high-water mark. Operation','allocator attribution is unavailable. Final memory/object/depth budgets return','to zero under the existing lifecycle oracle. No bounded-total-memory, cold I/O,','shared-link concurrency, native compatibility or scaling claim follows.','','Profiles include the whole fresh process and untimed work; cycles omit blocked','sleep. The zero L1 miss event on this guest must not be interpreted as a cache','benefit. Fixed and transfer sleeps are separate; OS oversleep affects observed','latency. The nominal pacing rate is not a measured achieved bandwidth.','']
    return '\n'.join(lines)
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--check',action='store_true');a=p.parse_args();value=render();path=ROOT/'measurements.md'
    if a.check:assert path.read_text()==value
    else:
        with path.open('x') as f:f.write(value)
    print('VALID')
