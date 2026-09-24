# Initial matrix-function release measurements

All 99 processes succeeded: 33 cases, three rounds, CPU 6, two warmups and
31 measured iterations per process. Each process used the harness's explicit
case-dependent repeat count. Full correctness validation ran once outside the
timer, including inverse validation. Timed work includes evaluation, checksum
consumption, retained-memory observation, and result drop; parsing and fixture
construction are excluded. RSS is whole-process maximum RSS and includes setup.

The table shows the range across rounds, with batch latency divided by the
reported repeat count. It is an initial absolute baseline, not a comparison
against the previous unsupported implementation. There are no deterministic
mismatches in result shape/checksum, work, allocations, requested/released
bytes, retained output bytes, or failure fields across the three rounds.

| Case | p50 µs/evaluation | p95 µs/evaluation | Work/evaluation | Allocations/evaluation | Retained output bytes | RSS KiB |
|---|---:|---:|---:|---:|---:|---:|
| mdeterm-2 | 1.92–1.94 | 1.95–1.96 | 39 | 12 | 0 | 2632–2892 |
| minverse-2 | 2.32–2.40 | 2.34–2.42 | 60 | 13 | 352 | 2600–2908 |
| mmult-2 | 2.81–2.94 | 2.99–3.03 | 57 | 15 | 352 | 2652–2896 |
| munit-2 | 1.11–1.14 | 1.13–1.15 | 16 | 8 | 352 | 2628–2924 |
| transpose-2 | 1.89–1.93 | 1.91–1.94 | 32 | 12 | 352 | 2780–2920 |
| mdeterm-8 | 10.58–10.63 | 10.71–11.06 | 538 | 20 | 0 | 2652–2920 |
| minverse-8 | 20.62–20.86 | 21.96–22.16 | 1515 | 21 | 5632 | 2652–2880 |
| mmult-8 | 20.79–20.83 | 22.18–22.41 | 1096 | 23 | 5632 | 2908–3160 |
| munit-8 | 2.25–2.26 | 2.28–2.32 | 76 | 8 | 5632 | 2660–2908 |
| transpose-8 | 8.27–8.64 | 8.33–8.70 | 327 | 20 | 5632 | 2648–2908 |
| mdeterm-16 | 42.11–43.50 | 43.15–46.69 | 2790 | 24 | 0 | 2908–3164 |
| minverse-16 | 114.87–116.04 | 118.35–119.27 | 10119 | 25 | 22528 | 2968–3148 |
| mmult-16 | 101.16–103.34 | 104.54–107.06 | 6565 | 27 | 22528 | 3036–3164 |
| munit-16 | 5.71–5.76 | 5.80–5.86 | 269 | 8 | 22528 | 2780–2844 |
| transpose-16 | 27.38–29.89 | 27.50–30.00 | 1444 | 24 | 22528 | 2876–3100 |
| transpose-64 | 577.10–594.77 | 586.56–600.60 | 27581 | 32 | 360448 | 4916–4956 |
| munit-64 | 73.30–75.29 | 76.76–79.79 | 4109 | 8 | 360448 | 3144–3372 |
| mmult-rect-4x8x2 | 7.47–7.74 | 7.56–7.82 | 280 | 20 | 704 | 2632–2780 |
| transpose-rect-2x8 | 3.56–3.59 | 3.62–3.63 | 87 | 16 | 1408 | 2632–2652 |
| minverse-singular-2 | 2.06–2.08 | 2.09–2.11 | 55 | 12 | 0 | 2780–2924 |
| mdeterm-nonsquare-2x1 | 1.40–1.44 | 1.41–1.46 | 22 | 10 | 0 | 2600–2664 |
| mmult-incompatible-2x3-2x2 | 2.89–2.97 | 3.05–3.16 | 57 | 15 | 0 | 2652–2920 |
| munit-zero | 0.89–0.91 | 0.90–0.91 | 12 | 7 | 0 | 2652–2920 |
| if-selected-mdeterm-8 | 10.79–11.56 | 10.85–11.65 | 549 | 20 | 0 | 2648–2804 |
| if-unselected-mdeterm-8 | 0.59–0.61 | 0.62–0.63 | 14 | 4 | 0 | 2624–2900 |
| if-selected-minverse-8 | 21.06–21.41 | 22.04–23.05 | 1526 | 21 | 5632 | 2652–2908 |
| if-unselected-minverse-8 | 0.59–0.60 | 0.61–0.62 | 14 | 4 | 0 | 2628–2872 |
| if-selected-mmult-8 | 20.90–21.06 | 22.23–22.82 | 1107 | 23 | 5632 | 3156–3164 |
| if-unselected-mmult-8 | 0.59–0.61 | 0.61–0.63 | 14 | 4 | 0 | 2652–2904 |
| if-selected-munit-8 | 2.48–2.49 | 2.51–2.54 | 87 | 8 | 5632 | 2648–2884 |
| if-unselected-munit-8 | 0.60–0.61 | 0.61–0.64 | 14 | 4 | 0 | 2644–2924 |
| if-selected-transpose-8 | 8.49–8.84 | 8.57–8.91 | 338 | 20 | 5632 | 2600–2844 |
| if-unselected-transpose-8 | 0.59–0.60 | 0.62–0.63 | 14 | 4 | 0 | 2652–2832 |

[Receipt](release-05/receipt.json) and [raw capture](release-05/capture.tar.gz)
retain the exact commands, child statuses, JSON rows, time output, build log,
frozen harness/lockfile, and runner. The source and full validation gates are
linked from the receipt. The exact release binary is retained outside tmpfs.
Loose capture files were removed after every archived file was verified.

Performance acceptance remains open. These short, shared-host measurements
do not establish regressions against existing scalar/reference workloads,
end-to-end spreadsheet speedups, large ill-conditioned numerical behavior,
or host-level recalculation throughput. Literal-only matrix branches reuse
completed results; reference-dependent branches remain uncached. Follow-up
comparisons must retain per-case results and apply the GOAL latency/RSS review
thresholds rather than averaging away regressions.
