"""Render all native metrics and separate observer totals."""
import driver as d

def render():
 a=d.read(d.P/'analysis.json');lines=['# 0843 retained diagnostic tables','','Native intervals are bootstrap intervals of six block values; time is ns, RSS KiB, faults counts.','','| Parts | Mode | Metric | Median | CI low | CI high |','|---:|---|---|---:|---:|---:|']
 for x in a['native_summary']:lines.append(f"| {x['parts']} | {x['mode']} | {x['metric']} | {x['median']:.3f} | {x['ci95_low']:.3f} | {x['ci95_high']:.3f} |")
 lines+=['','Callgrind guest instructions; each repeat is separate.','','| Parts | Repeat | Ir |','|---:|---:|---:|']
 for x in a['profiles']:lines.append(f"| {x['row']['parts']} | {x['row']['block']} | {x['profile']['summary_ir']} |")
 lines+=['','Heaptrack allocations whose full ancestry contains the owner; requested bytes, not peak live memory.','','| Parts | Repeat | Allocation events | Requested bytes |','|---:|---:|---:|---:|']
 for x in a['heaps']:
  o=x['attribution']['owner_attribution'];lines.append(f"| {x['row']['parts']} | {x['row']['block']} | {o['allocation_events']} | {o['requested_bytes']} |")
 return '\n'.join(lines)+'\n'
if __name__=='__main__':
 text=render();p=d.P/'tables.md'
 if p.exists():assert p.read_text()==text
 else:p.write_text(text)
 print('Tables PASS')
