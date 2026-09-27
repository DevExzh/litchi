# 0788 cached-Part memory attribution

This packet diagnoses the rejected 0787 cached-Part candidate. It does not
revise adoption thresholds or make an adoption decision; the exact candidate
is restored after capture. Native timing, whole-child process RSS and faults,
phase snapshots, and heaptrack allocation profiles measure different scopes.
Diagnostic probe measurements are never pooled with native timing or RSS.

- Reports/samples: 220 / 5092
- Native reports/samples: 120 / 3600
- Memory off/on samples: 496 / 496
- Heaptrack reports/samples: 16 / 480
- Bootstrap seed: 788078; endpoint indexes: [249, 9749]
- Previous 0787 rejection remains authoritative regardless of diagnostic counters.

| Shape | State | Floor | Workers | p50 | p95 | p99 | RSS | Minor faults | Major faults |
|---|---|---:|---:|---:|---:|---:|---:|---:|---:|
| small | fresh | 0 | 1 | 1.00265 | 1.00425 | 0.990094 | 0.99727 | 0.98613 | n/a |
| small | fresh | 0 | 4 | 1.00513 | 0.97257 | 1.01978 | 0.971716 | 0.998075 | n/a |
| small | fresh | 0 | 8 | 0.996991 | 0.970024 | 0.950055 | 1.00388 | 1.00654 | n/a |
| small | fresh | 0 | 32 | 0.996801 | 0.967332 | 0.959496 | 1.0594 | 1.00509 | n/a |
| small | primed | 0 | 1 | 1.00811 | 0.999977 | 0.996139 | 0.98018 | 1 | n/a |
| small | primed | 0 | 4 | 0.0331016 | 0.0332321 | 0.0405689 | 0.974631 | 0.997123 | n/a |
| small | primed | 0 | 8 | 0.0239119 | 0.0259797 | 0.0268818 | 0.987292 | 0.995114 | n/a |
| small | primed | 0 | 32 | 0.00971845 | 0.00986145 | 0.0100291 | 0.965783 | 0.574721 | n/a |
| large | primed | 0 | 4 | 0.0317175 | 0.0321964 | 0.0320613 | 0.998059 | 1.00776 | n/a |
| small | primed | 65536 | 4 | 1.00816 | 1.01952 | 1.019 | 0.968442 | 0.997452 | n/a |

Raw six-block distributions and paired deltas remain in paired.csv and analysis.json.
