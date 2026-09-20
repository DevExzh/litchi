#!/usr/bin/env python3
"""Derive scoped tables and retain all >5% review triggers."""
import hashlib,json,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
def read(name): return json.loads((P/name).read_text())
comparisons=read('native-comparisons.json');summary=read('native-summary.json')
flags=[dict(case=r['case'],phase=r['phase'],pair=r['pair'],metric=m,delta_ns=r['delta_ns'][m],delta_pct=v) for r in comparisons if r['kind']!='aa' for m,v in r['delta_pct'].items() if v>5]
(P/'native-review-triggers.json').write_text(json.dumps(flags,indent=2)+'\n')
noise={}
for phase in ['total_ns','capture_ns','clone_ns','settext_ns','commit_ns','apply_ns']:
 rows=[x for x in comparisons if x['phase']==phase and x['pair']=='a1/a0']
 noise[phase]={metric:dict(min_pct=min(x['delta_pct'][metric] for x in rows),max_pct=max(x['delta_pct'][metric] for x in rows),flags=[dict(case=x['case'],delta_pct=x['delta_pct'][metric]) for x in rows if abs(x['delta_pct'][metric])>5]) for metric in ['p50_ns','mean_ns','p95_ns','p99_ns']}
(P/'baseline-noise.json').write_text(json.dumps(noise,indent=2)+'\n')
lines=['| Workflow / input | Baseline medians (ms) | Candidate medians (ms) | Paired change |','| --- | ---: | ---: | ---: |']
for case in [r['case'] for r in read('cases.json')]:
 values={r['leg']:r['p50_ns']/1e6 for r in summary if r['case']==case and r['phase']=='total_ns'}
 pairs=[r['delta_pct']['p50_ns'] for r in comparisons if r['case']==case and r['phase']=='total_ns' and r['kind']!='aa']
 lines.append(f"| {case} | {values['a2']:.4f} / {values['a3']:.4f} | {values['b0']:.4f} / {values['b1']:.4f} | {pairs[0]:+.2f}% / {pairs[1]:+.2f}% |")
lines+=['','| Real one-edit phase | Baseline medians (ms) | Candidate medians (ms) |','| --- | ---: | ---: |']
for phase in ['capture_ns','clone_ns','settext_ns','commit_ns','apply_ns']:
 v={r['leg']:r['p50_ns']/1e6 for r in summary if r['case']=='one-real' and r['phase']==phase}
 lines.append(f"| {phase.removesuffix('_ns')} | {v['a2']:.4f} / {v['a3']:.4f} | {v['b0']:.4f} / {v['b1']:.4f} |")
lines+=['','| One-edit allocation diagnostic | Baseline | Candidate |','| --- | ---: | ---: |']
rs={r['run_phase']:r for r in read('allocation-summary.json') if r['case']=='one-real' and r['phase']=='total'}
for metric in ['alloc_calls','realloc_calls','requested_bytes']:
 lines.append(f"| {metric} | {rs['baseline']['metrics'][metric][0]:,} | {rs['candidate']['metrics'][metric][0]:,} |")
for metric in ['peak_above_start','net_live_change']:
 lines.append(f"| {metric} | {rs['baseline'][metric][0]:,} | {rs['candidate'][metric][0]:,} |")
(P/'tables.md').write_text('\n'.join(lines)+'\n')
profiles={}
for phase in ['baseline','candidate']:
 counters={}
 for count in [10,210]:
  counters[count]={}
  for line in (P/'profile'/phase/f'counters-{count}.txt').read_text().splitlines():
   parts=line.split('\t')
   if len(parts)>3 and parts[2] in ['cycles','instructions','branches','branch-misses','cache-misses','page-faults','task-clock']:
    counters[count][parts[2]]=float(parts[0])
 rss=next(int(x.split(':')[1]) for x in (P/'profile'/phase/'rss.txt').read_text().splitlines() if 'Maximum resident' in x)
 profiles[phase]=dict(stage="commit", counters=counters,
                       per_open_commit={k:(v-counters[10][k])/200 for k,v in counters[210].items()},
                       whole_child_peak_rss_kib=rss)
(P/'profile-summary.json').write_text(json.dumps(profiles,indent=2)+'\n')
print('tables and triggers derived; trigger count',len(flags))
