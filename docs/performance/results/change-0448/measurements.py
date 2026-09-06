#!/usr/bin/env python3
"""Render the entire delay-policy calibration with scoped model claims."""
import argparse,json
from pathlib import Path
ROOT=Path(__file__).resolve().parent
def render():
    v=json.loads((ROOT/'summary.json').read_text());lines=['# PPTX minimum-service calibration','','Both policies use 64 KiB ranges, 200 us fixed request delay and 25 MiB/s nominal','transfer targets. One build, CPU 2, one worker, two reversed repeats; 8 reports','and 240 retained samples (30 samples/3 warmups per process). These are delay-model','observations, not a production speedup or measured network bandwidth.','','| Corpus | Policy | Repeat | API p50 ms | p95 ms | p99 ms | p50 bootstrap 95% ms | Whole-process RSS MiB |','| --- | --- | --- | ---: | ---: | ---: | --- | ---: |']
    for row in v['rows']:
        t=row['timings']['api_sum_ns'];lo,hi=t['p50_bootstrap_95'];policy='minimum-service' if row['minimum_service'] else 'separate-sleeps'
        lines.append(f"| {row['corpus']} | {policy} | {row['repeat']} | {t['p50']/1e6:.6f} | {t['p95']/1e6:.6f} | {t['p99']/1e6:.6f} | {lo/1e6:.6f}–{hi/1e6:.6f} | {row['whole_process_peak_rss_bytes']/2**20:.3f} |")
    lines+=['','p50 uses the midpoint; p95/p99 use nearest rank. The 2,000 bootstrap resamples','describe within-process uncertainty, not machine/day uncertainty. All individual','phase timers and vectors remain in summary.json.','','| Corpus | Repeat | Minimum-service API p50 relative to separate sleeps |','| --- | --- | ---: |']
    for row in v['comparisons']:
        if row['metric']=='api_sum_ns.p50':lines.append(f"| {row['corpus']} | {row['repeat']} | {row['relative_percent']:+.3f}% |")
    lines+=['','| Corpus | Phase | Logical reads | Returned bytes | Nominal fixed + transfer floor ms |','| --- | --- | ---: | ---: | ---: |']
    for row in v['rows']:
        if row['repeat']=='R1' and row['minimum_service']:
            for phase,c in row['phase_read_deltas'].items():
                calls=c['logical_calls']['p50'];floor=calls*200000+c['transfer_delay_ns']['p50']
                lines.append(f"| {row['corpus']} | {phase} | {calls:,} | {c['returned_bytes']['p50']:,} | {floor/1e6:.6f} |")
    lines+=['','All underlying counters and nominal targets are equal across configurations,','repeats and samples. Every enclosing serial API clock satisfies the independently','checked combined service floor. The target is not a measured sleep duration.','Minimum-service credits time already spent, including wrapped-source work and','fixed-wait overshoot; separate sleeps request the full extra transfer wait.','']
    flags=[row for row in v['repeat_review'] if row['flagged']];lines.append(f'Absolute 5% repeat review: {len(flags)} flags.');lines.append('')
    for row in flags:lines.append(f"- {row['corpus']}, minimum-service={row['minimum_service']}, {row['metric']}: {row['relative_percent']:+.3f}% R2 versus R1.")
    lines+=['','Every paired difference remains in summary.json; all absolute 5% paired and','repeat triggers are reviewed in decision.json. No observations are discarded.','The sink retains full output; operation allocator attribution is unavailable.','Managed boundary gauges and whole-process RSS do not establish allocation peaks','or bounded total memory. Profiles include untimed work and omit blocked sleep.','No physical bandwidth, cold I/O, shared-link concurrency, native compatibility','or scaling claim follows.','']
    return '\n'.join(lines)
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--check',action='store_true');a=p.parse_args();value=render();path=ROOT/'measurements.md'
    if a.check:assert path.read_text()==value
    else:
        with path.open('x') as f:f.write(value)
    print('VALID')
