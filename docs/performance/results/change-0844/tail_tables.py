"""Render all confirmation metrics separately from original measurements."""
import driver as d

def render():
 a=d.read(d.P/'tail-analysis.json');lines=['# Separate 0844 tail confirmation','','All times are ns; RSS KiB; faults counts. Six paired blocks,1000samples perchild,3warmups.','','| Shape | Parts | Mode | Metric | Before median | After median | Paired ratio | CI low | CI high |','|---|---:|---|---|---:|---:|---:|---:|---:|']
 fmt=lambda x:'n/a' if x is None else f'{x:.6f}'
 for x in a['summary']:
  p=x['paired'];lines.append(f"| {x['shape']} | {x['parts']} | {x['mode']} | {x['metric']} | {x['before_median']:.3f} | {x['after_median']:.3f} | {fmt(p['estimate'])} | {fmt(p['ci95_low'])} | {fmt(p['ci95_high'])} |")
 return '\n'.join(lines)+'\n'
if __name__=='__main__':
 v=render();p=d.P/'tail-tables.md'
 if p.exists():assert p.read_text()==v
 else:p.write_text(v)
 print('Confirmation tables PASS')
