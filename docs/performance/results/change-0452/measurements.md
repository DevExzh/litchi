# Matched PPTX capture lifecycle measurements

Each row is a fresh CPU-2 process: 3 warmups and 30 samples. API time includes
open, planning and publication. RSS is whole-process high water, including
untimed corpus/gates/reporting; it is not a timed allocator peak.

| Provider | Corpus | Repeat | Baseline p50 ms | Candidate p50 ms | Change | Baseline RSS KiB | Candidate RSS KiB |
|---|---|---|---:|---:|---:|---:|---:|
| bytes | plain | R1 | 2.217 | 2.209 | -0.361% | 20096 | 20352 |
| bytes | plain | R2 | 2.216 | 2.221 | +0.217% | 19944 | 20108 |
| bytes | media-rich | R1 | 40.866 | 25.983 | -36.419% | 804152 | 802364 |
| bytes | media-rich | R2 | 28.326 | 26.087 | -7.903% | 802592 | 802676 |
| range | plain | R1 | 152.676 | 152.666 | -0.007% | 19912 | 20416 |
| range | plain | R2 | 152.660 | 152.683 | +0.015% | 20588 | 20676 |
| range | media-rich | R1 | 2572.729 | 1766.853 | -31.324% | 802432 | 803884 |
| range | media-rich | R2 | 2572.933 | 1774.087 | -31.048% | 802864 | 803688 |

Raw p95/p99, median bootstrap intervals, per-phase medians, all read-work
counters, reserved-memory gauges and every >5% paired/repeat flag are in
`measurements.json`. See `regression-review.md` for interpretation.

Range simulation uses 65,536-byte maximum returns, 200 microseconds per read,
and 25 MiB/s separate-sleep transfer pacing. It is not an actual network or
cold-filesystem measurement. Registry/default and native coverage are unchanged.
