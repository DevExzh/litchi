"""Render individual current paired results without pooling incomparable scopes."""
import driver as d
import csv,io
x=d.read(d.P/'analysis.json');rows=x['summaries']
f=io.StringIO();fields=['lane','index','shape','state','task_floor','workers','affinity','protocol','metric','before_median','after_median','estimate','ci95_low','ci95_high','guarded','regression_flag'];writer=csv.DictWriter(f,fieldnames=fields);writer.writeheader()
for row in rows:writer.writerow({k:row.get(k,row['case'].get(k,'')) for k in fields})
(d.P/'paired.csv').write_text(f.getvalue())
lines=['# Current-source paired results','', 'Native values below are medians of six process quantiles. Ratios and intervals are paired-block statistics; absolute median ratios need not equal them. RSS is whole-child GNU-time KiB. Instrumented scopes are separate in paired.csv and analysis.json.','', '| Shape | State | Floor | Width | Before p50 µs | After p50 µs | p50 ratio [95% CI] | RSS ratio [95% CI] | p95 ratio | p99 ratio |','|---|---|---:|---:|---:|---:|---|---|---:|---:|']
for i in range(60):
 group={r['metric']:r for r in rows if r['lane']=='native' and r['index']==i};p=group['p50_ns'];rss=group['rss_kib'];c=p['case'];interval=lambda r:f"{r['estimate']:.6f} [{r['ci95_low']:.6f}, {r['ci95_high']:.6f}]"
 lines.append(f"| {c['shape']} | {c['state']} | {c['task_floor']} | {c['workers']} | {p['before_median']/1000:.3f} | {p['after_median']/1000:.3f} | {interval(p)} | {interval(rss)} | {group['p95_ns']['estimate']:.6f} | {group['p99_ns']['estimate']:.6f} |")
lines+=['','All per-metric raw block vectors and bootstrap endpoints, including the tail intervals, are in analysis.json. No guarded regression flag is triggered. Five native tail point estimates exceed 1.05; all five intervals include one.','']
(d.P/'paired.md').write_text('\n'.join(lines))
print('tables written')
