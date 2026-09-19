#!/usr/bin/env python3
"""Keep all three allocation samples and explicit net releases per phase."""
import json
from pathlib import Path

P = Path(__file__).resolve().parent
out = []
for run in json.loads((P / 'allocation-runs.json').read_text()):
    lines = (P / run['output']).read_text().splitlines()
    header = next(x for x in lines if x.startswith('sample\t')).split('\t')
    rows = [dict(zip(header,map(int,x.split('\t')))) for x in lines if x.split('\t')[0].isdigit()]
    assert len(rows)==3
    for phase in ['capture','clone','settext','commit','apply','total']:
        metrics = {k.removeprefix(phase+'_'):[r[k] for r in rows]
                   for k in header if k.startswith(phase+'_') and not k.endswith('_ns')}
        out.append(dict(case=run['case'],phase=phase,metrics=metrics,
                        peak_above_start=[r[phase+'_peak_live_bytes']-r[phase+'_baseline_live_bytes'] for r in rows],
                        net_live_change=[r[phase+'_current_live_bytes']-r[phase+'_baseline_live_bytes'] for r in rows]))
(P / 'allocation-summary.json').write_text(json.dumps(out,indent=2)+'\n')
print('Allocation summary:',len(out),'case/phases, three samples each')
