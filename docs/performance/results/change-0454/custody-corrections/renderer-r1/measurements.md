# Change 0454 performance measurements

The provider lifecycle table retains the individual p50, p95 and p99 for every
phase and the complete API sum. The two external fixture rows are candidate-only
descriptive evidence: the preserved baseline typed-refuses the unnamed slide, so
no speedup against that refusal is claimed.

## Provider lifecycle controls

| Provider | Corpus | Repeat | Baseline API p50/p95/p99 ms | Candidate API p50/p95/p99 ms |
|---|---|---|---:|---:|
| bytes | media-rich | R1 | 25.297/25.450/25.549 | 25.378/25.504/25.541 |
| bytes | media-rich | R2 | 25.233/25.423/25.535 | 25.234/25.354/25.376 |
| bytes | plain | R1 | 2.230/2.248/2.261 | 2.258/2.272/2.274 |
| bytes | plain | R2 | 2.229/2.248/2.249 | 2.243/2.262/2.264 |
| range | media-rich | R1 | 1760.526/1765.391/1766.735 | 1760.373/1768.156/1775.334 |
| range | media-rich | R2 | 1770.601/1771.361/1776.913 | 1771.961/1782.907/1787.104 |
| range | plain | R1 | 152.797/153.496/154.412 | 152.799/152.871/153.501 |
| range | plain | R2 | 153.074/170.372/175.891 | 152.797/152.851/152.852 |

## External fixture (candidate-only)

| Provider | API p50/p95/p99 ms | RSS KiB | Logical read/work evidence |
|---|---:|---:|---|
| bytes | 1.857/1.891/1.896 | 13968 | planned source/destination calls 195/91 |
| range | 114.608/114.759/114.854 | 13668 | planned source/destination calls 265/116 |

## Scope and limits

- Every retained lane has 30 samples after 3 warmups in a fresh child, pinned to CPU 2 with one worker.
- Median intervals use 10,000 deterministic bootstrap resamples within each lane; p95 and p99 use linear interpolation over the retained samples.
- RSS and page-fault counters come from GNU `/usr/bin/time -v` around the whole child and include setup, fixture construction and retained samples.
- Logical read counters and budget work are copied from the Rust reports and checked for per-lane stability; they do not assert physical device I/O.
- Allocation calls/bytes, live bytes and allocator-region peaks are unavailable from these ordinary binaries. RSS is not an allocation substitute.
- Range lanes use a 64 KiB logical cap, 200 microseconds per request and 25 MiB/s separate sleeps. The external range lane has its own 256-byte/100 microsecond logical adapter settings.
- Review flags above 5% are retained for individual inspection; no flag is silently converted into an acceptance claim.
