# PPTX shared decoded payload measurements

Normal latency and allocator diagnostics use separate binaries and processes.
Each row has 3 warmups / 30 samples. API includes open, plan and publication.

| Provider | Corpus | Repeat | Baseline API p50 ms | Candidate API p50 ms | Change |
|---|---|---|---:|---:|---:|
| bytes | plain | R1 | 2.215 | 2.217 | +0.053% |
| bytes | media-rich | R1 | 26.122 | 25.340 | -2.995% |
| range | plain | R1 | 152.789 | 152.793 | +0.003% |
| range | media-rich | R1 | 1761.152 | 1760.201 | -0.054% |
| range | media-rich | R2 | 1761.330 | 1760.165 | -0.066% |
| range | plain | R2 | 152.685 | 152.803 | +0.077% |
| bytes | media-rich | R2 | 26.123 | 25.282 | -3.217% |
| bytes | plain | R2 | 2.223 | 2.230 | +0.296% |

| Allocator corpus | Repeat | Plan allocated bytes before / after | Plan live growth before / after | Publication region peak before / after |
|---|---|---:|---:|---:|
| plain | R1 | 982720.0 / 982720.0 | 56283.0 / 56283.0 | 1417731.0 / 1417729.0 |
| media-rich | R1 | 51760790.0 / 34983382.0 | 50425902.0 / 33648494.0 | 238097429.0 / 221320019.0 |
| media-rich | R2 | 51760790.0 / 34983382.0 | 50425902.0 / 33648494.0 | 238097429.0 / 221320019.0 |
| plain | R2 | 982720.0 / 982720.0 | 56283.0 / 56283.0 | 1417731.0 / 1417729.0 |

Raw samples and measurements.json retain phase medians, p95/p99, median intervals,
allocator counters/peaks, process RSS, source work and owner read/budget gauges.
Process RSS includes untimed fixture construction. Allocator region peaks are
absolute live bytes including region entry, not physical RSS. Conservative staging
admission remains. Range is 64 KiB/200 us/25 MiB/s separate-sleep simulation.
All >5% paired and repeat flags are reviewed individually; no native/cold/scaling
coverage is promoted.
