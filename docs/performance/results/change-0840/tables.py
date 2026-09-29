"""Render the complete frozen summaries, without changing raw evidence."""
import driver as d
import csv,io,math
x=d.read(d.P/'analysis.json');rows=x['summary']
fields=['lane','case','metric','before_median','after_median','estimate','ci95_low','ci95_high','guarded','regression_flag']
out=io.StringIO();writer=csv.DictWriter(out,fieldnames=fields,lineterminator='\n',extrasaction='ignore');writer.writeheader();writer.writerows(rows)
(d.P/'paired.csv').write_text(out.getvalue())
lines=['# Fresh CFB emission paired results','',f"Decision: **{x['decision']}**. {x['reports']} reports / {x['samples']} samples; {len(x['regression_flags'])} frozen guard flags.",'','Ratios are medians of six matched process ratios. Absolute columns are medians of process summaries; they need not divide to the paired ratio.','', '| Case | Before p50 ms | After p50 ms | p50 ratio [95% CI] | RSS ratio [95% CI] | p95 ratio | p99 ratio |','| --- | ---: | ---: | --- | --- | ---: | ---: |']
def find(case,metric,lane='native'):return next(r for r in rows if r['case']==case and r['metric']==metric and r['lane']==lane)
def interval(r):return f"{r['estimate']:.6f} [{r['ci95_low']:.6f}, {r['ci95_high']:.6f}]"
for case in d.read(d.P/'plan.json')['cases']:
 p=find(case,'p50_ns');rss=find(case,'rss_kib')
 lines.append(f"| {case} | {p['before_median']/1e6:.6f} | {p['after_median']/1e6:.6f} | {interval(p)} | {interval(rss)} | {find(case,'p95_ns')['estimate']:.6f} | {find(case,'p99_ns')['estimate']:.6f} |")
lines+=['','## Operation allocations','','| Case | Calls before → after | Requested bytes before → after | Peak above entry before → after | Retained delta before → after |','| --- | ---: | ---: | ---: | ---: |']
for case in d.read(d.P/'plan.json')['cases']:
 values=[find(case,m,'allocation') for m in ['allocation_calls','allocated_bytes','region_peak_live_bytes_minus_entry','retained_live_bytes_delta']]
 lines.append('| '+case+' | '+' | '.join(f"{r['before_median']:g} → {r['after_median']:g}" for r in values)+' |')
lines+=['','Full intervals and all fault/quantile rows are in paired.csv and analysis.json. RSS is whole-child GNU-time maximum, including setup and verification. Allocator regions exclude setup, verification and output destruction.','']
(d.P/'paired.md').write_text('\n'.join(lines))
print('tables PASS',len(rows),'rows')
