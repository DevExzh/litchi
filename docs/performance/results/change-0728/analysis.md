# Current DOC/PPT baseline

Ranges below span six independent native processes per route. Timing is microseconds; phase percentages are per-process medians of within-sample ratios. Allocation counters span three separate instrumented processes. No before/after speedup, RSS, cold-cache, or concurrent claim follows.

| Case | Route | Whole p50 μs range | Whole mean μs range | Stage % range | Finish % range |
| --- | --- | ---: | ---: | ---: | ---: |
| docfloat | format default | 984.82–1089.01 | 991.14–1078.85 | — | — |
| docfloat | container reuse | 243.05–247.64 | 243.66–247.45 | 53.83–54.10 | 30.13–30.67 |
| docfloat | container rewrite | 166.80–177.67 | 169.68–178.17 | 53.80–54.22 | 23.02–24.94 |
| docnohf | format default | 108.01–110.83 | 110.43–112.89 | — | — |
| docnohf | container reuse | 43.88–44.80 | 44.60–45.30 | 56.52–57.42 | 28.95–29.29 |
| docnohf | container rewrite | 32.98–33.68 | 33.17–34.35 | 56.06–56.84 | 24.50–24.97 |
| ppt45543 | format default | 1128.81–1140.66 | 1115.55–1132.03 | — | — |
| ppt45543 | container reuse | 271.53–278.34 | 278.43–286.68 | 43.58–43.77 | 17.96–18.78 |
| ppt45543 | container rewrite | 210.36–211.63 | 207.06–209.15 | 41.66–41.83 | 9.61–9.74 |

The common-container PPT route is an alternative control, not a decomposition of its public save. DOC also runs in separate processes: subtracting these route medians would not estimate a nested phase. Staging includes render, reopen, recapture and discovery. Allocation regions retain returned editors/output vectors at the recorded boundary; peak live values are region increments, not whole-process RSS.

See analysis.json for all sample-derived per-process p95, p99, maxima and phase fractions, plus exact allocation values.
