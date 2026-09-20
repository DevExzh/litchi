#!/usr/bin/env python3
"""Render every paired observation from the verified analyses, without pooling."""
import json
from pathlib import Path
P=Path(__file__).resolve().parent
primary=json.loads((P/'analysis.json').read_text());reads=json.loads((P/'read-controls-analysis.json').read_text())
lines=['# 0721 paired measurements','', 'All deltas are candidate / baseline minus one. Negative latency deltas are faster. Times are microseconds. No pair or sample is omitted.','', '## Primary native','', '| Pair | Corpus | Phase | Baseline p50 | Candidate p50 | p50 delta | Mean delta | p95 delta | p99 delta |','| --- | --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |']
for pair,data in primary['paired_comparisons'].items():
 for corpus,phases in data['corpora'].items():
  for phase,row in phases.items():
   n=row['native'];lines.append(f"| {pair} | {corpus} | {phase} | {n['p50']['baseline']/1000:.2f} | {n['p50']['candidate']/1000:.2f} | {n['p50']['delta_percent']:+.2f}% | {n['mean']['delta_percent']:+.2f}% | {n['p95']['delta_percent']:+.2f}% | {n['p99']['delta_percent']:+.2f}% |")
lines+=['','## Read controls','','| Pair | Control | Baseline p50 | Candidate p50 | p50 delta | Mean delta | p95 delta | p99 delta |','| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |']
for pair,data in reads['paired_comparisons'].items():
 for control,row in data.items():
  n=row['native'];lines.append(f"| {pair} | {control} | {n['p50']['baseline']/1000:.2f} | {n['p50']['candidate']/1000:.2f} | {n['p50']['delta_percent']:+.2f}% | {n['mean']['delta_percent']:+.2f}% | {n['p95']['delta_percent']:+.2f}% | {n['p99']['delta_percent']:+.2f}% |")
lines+=['','## Allocator requests','','Instrumented request counts and requested bytes are separate from native latency. Values below are medians of three samples.','','| Pair | Corpus | Phase | Requests baseline → candidate | Requested bytes baseline → candidate |','| --- | --- | --- | ---: | ---: |']
for pair,data in primary['paired_comparisons'].items():
 for corpus,phases in data['corpora'].items():
  for phase,row in phases.items():
   if 'allocator' not in row:continue
   a=row['allocator']['allocation'];c=a['request_count']['p50'];b=a['requested_bytes']['p50'];lines.append(f"| {pair} | {corpus} | {phase} | {c['baseline']:.0f} → {c['candidate']:.0f} | {b['baseline']:.0f} → {b['candidate']:.0f} |")
lines+=['','## Failed hard gates','']
for lane,value in [('primary',primary),('read',reads)]:
 failed=[r for r in value['decision']['hard_gates'] if not r['pass']]
 lines.append(f'{lane}: {len(failed)} failed of {len(value["decision"]["hard_gates"])}.')
 for row in failed:lines.append('- `'+json.dumps(row,sort_keys=True)+'`')
 lines.append('')
(P/'measurements.md').write_text('\n'.join(lines).rstrip()+'\n')
print('Rendered all paired native, read and allocator observations')
