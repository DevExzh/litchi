# Reference-metadata performance capture history

The final authorized revised capture retained 840 baseline rows and 4,080 candidate rows (4,920 timed rows) in session `38187`, with three warmups and fifteen samples per phase and case. The source freeze hash is `b4a0f0c9665e10f67fafa1b179bfbfdd8be2c65e6ec897bac86db1b2afc9330a`.

| session | disposition | baseline rows | candidate rows | timed rows |
| --- | --- | ---: | ---: | ---: |
| `38187` | complete authorized revised frozen-source pair | 840 | 4,080 | 4,920 |

The setup-only launch failure is retained under [`diagnostic-revised-freeze-path-failure/`](diagnostic-revised-freeze-path-failure/); it produced zero build, preflight, and timed rows. The stale retained verifier bound is preserved under [`diagnostic-stale-read-bound-verifier/`](diagnostic-stale-read-bound-verifier/); it rejects the six exact two-read computed ROW/COLUMN lanes because its generic metadata rule still expects zero reads. The independent root audit [`root-performance-audit.json`](../../root-performance-audit.json) applied the corrected read contract and passed all 4,920 retained rows.

The superseded complete 99-case capture remains at [`diagnostics/computed-array-preflight/performance-results/`](../../diagnostics/computed-array-preflight/performance-results/), including its +5.865% SUMIFS evaluate observation. The two complete captures contain 9,570 timed rows in total. The revised analysis and p50 tables are [`capture-analysis.md`](capture-analysis.md) and [`performance-report.md`](performance-report.md).
