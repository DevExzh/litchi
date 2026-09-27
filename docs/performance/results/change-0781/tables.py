"""Render retained 0781 report CSVs; --check verifies exact bytes."""
import csv
import io
import json
from pathlib import Path
import statistics
import sys

P=Path(__file__).resolve().parent
A=json.loads((P/'analysis.json').read_text())

def emit(name, rows):
    out=io.StringIO(newline='');w=csv.writer(out,lineterminator='\n');w.writerows(rows);s=out.getvalue()
    if '--check' in sys.argv:assert (P/name).read_text()==s,name
    else:(P/name).write_text(s)

native=[['lane','shape','mode','leg','metric','unit','median_of_process_values','min','max','spread_percent','flag_over_5_percent']]
processes=[['lane','shape','mode','leg','block','mean_ns','p50_ns','p95_ns','p99_ns','rss_kib','report']]
pairs=[['lane','shape_mode','metric','block','before','after','change_percent','over_5_percent']]
alloc=[['lane','shape','mode','leg','metric','block0_p50','block1_p50','spread_percent','flag_over_5_percent']]
flags=[['lane','kind','flag_json']]
for lane,section in [('primary',A)]:
    n=section['native']['analysis'];al=section['allocation']['analysis']
    for g in n['groups'].values():
        base=[lane,g['shape'],g['mode'],g['leg']]
        metrics=dict(g['elapsed_distribution_across_processes']);metrics['rss_kib']=g['per_process_rss_distribution']
        for metric,v in metrics.items():native.append(base+[metric,'KiB' if metric=='rss_kib' else 'ns',statistics.median(v['values']),min(v['values']),max(v['values']),v['spread_percent'],v['flag_over_5_percent']])
        for v in g['processes']:processes.append(base+[v['block']]+[v['stats'][k] for k in ['mean','p50','p95','p99']]+[v['rss_kib'],v['report']])
    for key,g in n['paired_by_block_before_after'].items():
        for metric,v in g['metrics'].items():
            for row in v['by_block']:pairs.append([lane,key,metric]+[row[k] for k in ['block','before','after','change_percent','over_5_percent']])
    for g in al['groups'].values():
        for metric,v in g['metrics'].items():alloc.append([lane,g['shape'],g['mode'],g['leg'],metric]+v['repeat_p50_values']+[v['spread_percent'],v['flag_over_5_percent']])
    for kind,values in [('native_spread',n['spread_flags_over_5_percent']),('native_paired',n['regression_flags_over_5_percent']),('allocation_spread',al['spread_flags_over_5_percent']),('allocation_paired',al['regression_flags_over_5_percent'])]:
        for v in values:flags.append([lane,kind,json.dumps(v,sort_keys=True,separators=(',',':'))])
for name,rows in [('native-summary.csv',native),('native-processes.csv',processes),('native-pairs.csv',pairs),('allocation-summary.csv',alloc),('all-flags.csv',flags)]:emit(name,rows)
print('Five tables match analysis.' if '--check' in sys.argv else 'Wrote five tables.')
