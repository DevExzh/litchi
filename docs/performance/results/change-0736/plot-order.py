#!/usr/bin/env python3
"""Order-preserving view: pointwise process medians, not confidence bands."""
import json
from pathlib import Path
import statistics
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

P = Path(__file__).resolve().parent
OLD = P.parent / 'change-0735'
manifest = json.loads((OLD / 'captures/manifest.json').read_text())
fig, axes = plt.subplots(1, 2, figsize=(11, 4), layout='constrained')
for ax, case, title in zip(axes, ['primary', 'secondary'], ['45543.ppt', '41246-1.ppt']):
    for variant, color in [('baseline', '#2665a8'), ('candidate', '#c45423')]:
        reports = [json.loads((OLD / 'captures' / r['output']).read_text()) for r in manifest['runs']
                   if r['lane'] == 'native' and r['case'] == case and r['variant'] == variant]
        assert len(reports) == 9
        values = [[s['phase_ns']['whole_ns'] / 1000 for s in r['samples']] for r in reports]
        for row in values:
            ax.plot(range(50), row, color=color, alpha=.12, linewidth=.6)
        ax.plot(range(50), [statistics.median(col) for col in zip(*values)], color=color,
                label=f'{variant}: pointwise median of 9 processes', linewidth=1.6)
    ax.set_title(title)
    ax.set_xlabel('Recorded sample index (after 3 warmups)')
    ax.set_ylabel('Public owner time (µs)')
    ax.grid(alpha=.2)
    ax.legend(fontsize=7)
fig.suptitle('0735 retained sample order; thin lines show every process')
fig.savefig(P / 'sample-order.png', dpi=160)
plt.close(fig)
(P / 'plot-receipt.json').write_text(json.dumps(dict(matplotlib=matplotlib.__version__,
    samples=1800, scope='Descriptive trajectories; no independent-sample or confidence-band claim.'), indent=2) + '\n')
