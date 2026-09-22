#!/usr/bin/env python3
"""Render all paired flags and compact ranges from the audited result."""
import json
from pathlib import Path
P=Path(__file__).resolve().parent
x=json.loads((P/'analysis.json').read_text());out=[]
for case in ['primary','secondary']:
 out.append('\n'+case)
 for v in ['baseline','candidate']:
  rows=[r['timing'] for r in x['native_processes'] if r['case']==case and r['variant']==v]
  out.append(v+' '+str({k:[min(r[k] for r in rows)/1000,max(r[k] for r in rows)/1000] for k in ['p50','mean','p95','p99','maximum']}))
for r in x['native_pairs']:
 flagged={k:round(v['percent'],6) for k,v in r['differences'].items() if v['flag_gt_5pct']}
 out.append(str((r['case'],r['cycle'],r['repeat'],flagged)))
(P/'report-stats.txt').write_text('\n'.join(out)+'\n');print('\n'.join(out))
