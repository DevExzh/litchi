"""Render complete individual native metrics and scoped guest instruction rows."""
import driver as d
import csv,io
v=d.read(d.P/'analysis.json')
def publish(name,text):
 p=d.P/name
 if p.exists():assert p.read_text()==text
 else:p.write_text(text)
rows=[]
for x in v['summary']:
 p=x['paired'];rows.append(dict(shape=x['shape'],mode=x['mode'],metric=x['metric'],before=x['before_median'],after=x['after_median'],ratio=p['estimate'],low=p['ci95_low'],high=p['ci95_high']))
f=io.StringIO(newline='');w=csv.DictWriter(f,fieldnames=list(rows[0]),lineterminator='\n');w.writeheader();w.writerows(rows);publish('native.csv',f.getvalue())
fmt=lambda x:'n/a' if x is None else f'{x:.6f}'
lines=['# Current-source native diagnostic','', 'Absolute values are medians of six process values. Paired estimates are medians of matched block ratios, not quotients of the displayed absolute medians. Whole-child RSS includes setup/readback. Fault counters are whole-child observations.','', '| Shape | Mode | Metric | Before | After | Paired ratio | 95% lower | 95% upper |','| --- | --- | --- | ---: | ---: | ---: | ---: | ---: |']
for x in rows:lines.append('| '+' | '.join(str(x[k]) if k in ['shape','mode','metric','before','after'] else fmt(x[k]) for k in x)+' |')
lines+=['','## Threshold observations','',f"{len(v['regression_flags'])} native p50/RSS guard flags; {len(v['spread_flags'])} max/min block-p50 spread flags. Diagnostic only; no adoption is authorized.",'']
publish('native.md','\n'.join(lines))
lines=['# Operation-scoped Callgrind diagnostic','', 'Guest Ir is not native cycles, latency, physical copies or a causal share. Each row is one operation in a fresh process. No profile confidence interval is inferred from two repeats.','', '| Shape | Repeat | Before Ir | After Ir | Ratio |','| --- | ---: | ---: | ---: | ---: |']
for x in v['profile_pairs']:lines.append(f"| {x['shape']} | {x['block']} | {x['before_ir']} | {x['after_ir']} | {x['ratio']:.8f} |")
for x in v['profile_pairs']:
 lines+=['',f"## {x['shape']} repeat {x['block']}: largest recorded self-cost differences",'', 'A missing separately emitted function is not zero logical work; names can move into callers through inlining. These rows partition recorded self costs; descendant inclusive costs are not added.','', '| Function | Before self Ir | After self Ir | Delta |','| --- | ---: | ---: | ---: |']
 for y in x['self_differences'][:20]:lines.append(f"| `{y['name'].replace('|','&#124;')}` | {y['before_recorded_self_ir']} | {y['after_recorded_self_ir']} | {y['delta']} |")
publish('profiles.md','\n'.join(lines)+'\n')
print('Tables PASS',len(rows),'native metric rows',len(v['profile_pairs']),'profile pairs')
