"""Render report tables from admitted analysis; --check replays without writes."""
import csv
import io
import json
from pathlib import Path
import sys

P = Path(__file__).resolve().parent
A = json.loads((P / 'analysis.json').read_text())

def emit(name, rows):
    out = io.StringIO(newline='')
    writer = csv.writer(out, lineterminator='\n')
    writer.writerows(rows)
    value = out.getvalue()
    if '--check' in sys.argv:
        assert (P / name).read_text() == value, name
    else:
        (P / name).write_text(value)

native = [['corpus', 'phase', 'policy', 'metric', 'unit', 'median', 'min', 'max', 'spread_percent', 'flag_over_5_percent']]
processes = [['corpus', 'phase', 'policy', 'block', 'count', 'mean_ns', 'p50_ns', 'p95_ns', 'p99_ns', 'rss_kib', 'report']]
for key, g in A['native']['analysis']['groups'].items():
    prefix = [g['corpus_id'], g['phase'], g['policy']]
    metrics = dict(g['elapsed_distribution_across_4_processes'])
    metrics['rss'] = g['whole_process_rss_distribution_across_4_processes']
    for metric, v in metrics.items():
        native.append(prefix + [metric, 'KiB' if metric == 'rss' else 'ns'] + [v[k] for k in ['median','min','max','spread_percent','flag_over_5_percent']])
    for v in g['processes']:
        processes.append(prefix + [v[k] for k in ['block','count','mean','p50','p95','p99','rss_kib','report']])
emit('native-summary.csv', native)
emit('native-processes.csv', processes)
emit('spread-flags.csv', [['corpus','phase','policy','metric','spread_percent']] + [v['group'] + [v['metric'],v['spread_percent']] for v in A['native']['analysis']['spread_flags_over_5_percent']])
aa = [['corpus','phase','metric','block_0_percent','block_1_percent','block_2_percent','block_3_percent','flag_over_5_percent']]
for g in A['native']['analysis']['default_vs_full_aa_like_controls']:
    for metric,v in g['metrics'].items():
        aa.append([g['corpus_id'],g['phase'],metric] + v['change_percent_by_block'] + [v['flag_over_5_percent']])
emit('default-full-controls.csv', aa)
alloc = [['corpus','phase','policy','metric','block_0_p50','block_1_p50','spread_percent','flag_over_5_percent']]
for g in A['allocation']['analysis']['groups'].values():
    for metric,v in g['metrics'].items():
        alloc.append([g['corpus_id'],g['phase'],g['policy'],metric] + v['repeat_p50_values'] + [v['spread_percent'],v['flag_over_5_percent']])
emit('allocation-summary.csv', alloc)
counters = [['corpus','policy','repeat','event','samples_3','samples_23','marginal_per_extra_sample','runtime_percent_samples_3','runtime_percent_samples_23']]
for g in A['counter_diagnostics']['marginal_pairs']:
    for event,v in g['events'].items():
        counters.append([g['corpus_id'],g['policy'],g['repeat'],event] + [v[k] for k in counters[0][4:]])
emit('counter-summary.csv', counters)
print('Six report tables match analysis.' if '--check' in sys.argv else 'Wrote six report tables.')
