"""Expose every metric and allocation guard in readable units."""
import driver as d

def render():
 a=d.read(d.P/'analysis.json');lines=['# 0844 complete comparison tables','','Times are ns; RSS KiB; faults counts. Ratios are medians of six matched block ratios.','','| Shape | Parts | Mode | Metric | Before median | After median | Ratio | CI low | CI high |','|---|---:|---|---|---:|---:|---:|---:|---:|']
 fmt=lambda x:'n/a' if x is None else f'{x:.6f}'
 for x in a['summary']:
  p=x['paired'];lines.append(f"| {x['shape']} | {x['parts']} | {x['mode']} | {x['metric']} | {x['before_median']:.3f} | {x['after_median']:.3f} | {fmt(p['estimate'])} | {fmt(p['ci95_low'])} | {fmt(p['ci95_high'])} |")
 lines+=['','Heaptrack owner-ancestry allocation events/requested bytes. Both repeats shown; no peak-live-memory claim.','','| Shape | Parts | Metric | Before repeats | After repeats | Ratio |','|---|---:|---|---|---|---:|']
 for x in a['allocations']:lines.append(f"| {x['shape']} | {x['parts']} | {x['metric']} | {x['before']} | {x['after']} | {x['ratio']:.6f} |")
 return '\n'.join(lines)+'\n'
if __name__=='__main__':
 s=render();p=d.P/'tables.md'
 if p.exists():assert p.read_text()==s
 else:p.write_text(s)
 print('Tables PASS')
