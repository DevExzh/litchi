#!/usr/bin/env python3
"""Plot all native processes and pointwise medians; phases remain separate."""
import json
from pathlib import Path
import statistics
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

HERE = Path(__file__).resolve().parent
fig, axes = plt.subplots(2, 2, figsize=(14, 8), sharex=True)
for row, (name, source) in enumerate([
    ('Sample fields', HERE / 'analysis.json'),
    ('Startup arguments', HERE / 'argv-control/analysis.json'),
]):
    data = json.loads(source.read_text())
    for col, case in enumerate(['primary', 'secondary']):
        ax = axes[row][col]
        rows = [p for p in data['processes'] if p['lane'] == 'native' and p['case'] == case]
        arms = sorted({p['arm'] for p in rows})
        for color, arm in zip(['#0072B2', '#D55E00', '#009E73', '#CC79A7'], arms):
            series = [[v / 1e6 for v in p['times_ns']] for p in rows if p['arm'] == arm]
            assert len(series) == 9 and all(len(v) == 50 for v in series)
            for values in series:
                ax.plot(range(1, 51), values, color=color, alpha=.10, linewidth=.6)
            ax.plot(range(1, 51), [statistics.median(v) for v in zip(*series)],
                    color=color, linewidth=1.2, label=arm)
        ax.set_title(f'{name}: {case}')
        ax.set_ylabel('Owner duration (ms)')
        ax.set_xlabel('Measured sample position')
        ax.grid(alpha=.2)
        ax.legend(fontsize=8, ncol=2)
fig.suptitle('0738: all 144 native processes; bold lines are pointwise medians of 9 runs')
fig.tight_layout()
fig.savefig(HERE / 'sample-position.png', dpi=160)
