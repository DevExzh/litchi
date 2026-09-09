#!/usr/bin/env python3
"""Render the separately traced operations; these are not formal latency samples."""
import statistics

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

import sync_profiles as profiles


def main():
    summary = profiles.analyze()
    if summary != profiles.read(profiles.ROOT / 'sync-summary.json'):
        raise ValueError('sync summary changed')
    rows = [r for r in summary['rows'] if 'paired_samples' in r]
    figure, axes = plt.subplots(1, 2, figsize=(12, 4.8), constrained_layout=True)
    colors = {'before': '#b65306', 'after': '#008a99'}
    labels = []
    data = []
    for index, row in enumerate(rows):
        run = row['run']
        samples = row['paired_samples']
        elapsed = [s['elapsed_ns'] / 1e6 for s in samples]
        sync = [s['fdatasync_ns'] / 1e6 for s in samples]
        labels.append(f"{run['phase'].title()} R{run['repeat']}")
        axes[0].bar(index, statistics.median(elapsed), color='#c9d0d4', width=.65)
        axes[0].bar(index, statistics.median(sync), color=colors[run['phase']], width=.65)
        axes[0].text(index, statistics.median(elapsed) + .09,
                     f"{statistics.median(s['sync_fraction'] for s in samples):.1%}", ha='center', fontsize=10)
        axes[1].scatter(sync, elapsed, color=colors[run['phase']], alpha=.65,
                        marker='o' if run['repeat'] == 1 else '^', label=labels[-1], s=25)
        data.append({'label': labels[-1], 'elapsed_ms': elapsed, 'fdatasync_ms': sync,
                     'median_sync_fraction': statistics.median(s['sync_fraction'] for s in samples)})
    axes[0].set_xticks(range(len(rows)), labels)
    axes[0].set_ylabel('Traced operation median (ms)')
    axes[0].set_ylim(0, 4.7)
    axes[0].set_title('Colored bar: median fdatasync duration\nGray bar: median whole operation')
    axes[1].set_xlabel('fdatasync duration (ms)')
    axes[1].set_ylabel('Containing traced operation (ms)')
    axes[1].set_title('One point per matched traced sample')
    axes[1].legend(fontsize=9)
    for axis in axes:
        axis.spines[['top', 'right']].set_visible(False)
        axis.grid(axis='y', alpha=.15)
        axis.set_axisbelow(True)
    figure.suptitle('DOCX file replay: synchronization attribution\n64 source / 64 authored paragraphs; data-sync policy retained', fontsize=14)
    figure.savefig(profiles.ROOT / 'sync-attribution.png', dpi=150)
    figure.savefig(profiles.ROOT / 'sync-attribution.svg')
    plt.close(figure)
    profiles.write(profiles.ROOT / 'sync-plot-data.json', {
        'source': profiles.meta(profiles.ROOT / 'sync-summary.json'),
        'renderer': profiles.meta(__file__), 'rows': data,
        'scope': '30 straced samples per child; three warmups and one preflight excluded. Tracing perturbs timing; no untraced latency claim.'})


if __name__ == '__main__':
    main()
