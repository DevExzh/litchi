#!/usr/bin/env python3
"""Show every 50-sample process, retaining chronological sample order."""
import json
from pathlib import Path
import statistics
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

P = Path(__file__).resolve().parent
data = json.loads((P/'analysis.json').read_text())
fig, axes = plt.subplots(2,2,figsize=(12,7),layout='constrained')
for col,case in enumerate(('primary','secondary')):
    for row,arms in enumerate((('archive','legacy-a','legacy-b'),('retained','drained'))):
        ax = axes[row,col]
        for arm,color in zip(arms,('#2665a8','#c45423','#5c8c4a')):
            runs = [p['times_ns'] for p in data['processes']
                    if p['lane']=='native' and p['case']==case and p['arm']==arm]
            assert len(runs)==9 and all(len(r)==50 for r in runs)
            for run in runs:
                ax.plot(range(50),[n/1000 for n in run],color=color,alpha=.12,linewidth=.6)
            ax.plot(range(50),[statistics.median(v)/1000 for v in zip(*runs)],color=color,label=arm)
        ax.set_title(case + (' — legacy controls' if row==0 else ' — matched strict controls'))
        ax.set_xlabel('Sample index after three warmups')
        ax.set_ylabel('Public owner time (µs)')
        ax.grid(alpha=.2)
        ax.legend(fontsize=8)
fig.suptitle('Unchanged PPT owner: harness lifecycle sensitivity\nThin lines: all processes; thick lines: pointwise medians, not confidence bands')
fig.savefig(P/'sample-order.png',dpi=160)
plt.close(fig)
(P/'plot-receipt.json').write_text(json.dumps(dict(matplotlib=matplotlib.__version__,
    processes=90,samples=4500,scope='All 50-sample processes; 18 fresh one-sample controls reported separately.'),indent=2)+'\n')
