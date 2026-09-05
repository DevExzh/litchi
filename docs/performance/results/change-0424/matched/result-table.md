# 0424 matched source-backed lifecycle baseline

Rows preserve full vectors in `summary.json`. Normal elapsed and RSS comparisons are diagnostic; allocator comparisons cover resource vectors only. No result authorizes a performance claim.

| Lane | Corpus | Role | Repeat | Samples | p50 ns | Mean ns | p95 ns | p99 ns | RSS KiB | Alloc calls mean | Allocated bytes mean | Region peak mean |
|---|---|---|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| normal | plain | control | R1 | 100 | 2161228 | 2163919.290 | 2180938 | 2188508 | 82,732 | — | — | — |
| normal | plain | candidate | R1 | 100 | 2178628 | 2182599.280 | 2195958 | 2219098 | 82,672 | — | — | — |
| normal | plain | candidate | R2 | 100 | 2182538 | 2185961.190 | 2199568 | 2205639 | 82,716 | — | — | — |
| normal | plain | control | R2 | 100 | 2166192 | 2168673.200 | 2181218 | 2193378 | 82,724 | — | — | — |
| normal | media_rich | control | R1 | 100 | 261755724 | 261774154.230 | 262205940 | 262271238 | 803,692 | — | — | — |
| normal | media_rich | candidate | R1 | 100 | 261721836 | 261752279.090 | 262142344 | 262311853 | 803,452 | — | — | — |
| normal | media_rich | candidate | R2 | 100 | 261876991 | 261909518.500 | 262344720 | 262457560 | 803,688 | — | — | — |
| normal | media_rich | control | R2 | 100 | 272291187 | 272293936.450 | 272522201 | 272684594 | 805,224 | — | — | — |
| allocator | plain | control | R1 | 30 | — | — | — | — | 82,620 | 10993.000 | 8916495.000 | 1059128.000 |
| allocator | plain | candidate | R1 | 30 | — | — | — | — | 82,708 | 10993.000 | 8916495.000 | 1059132.000 |
| allocator | plain | candidate | R2 | 30 | — | — | — | — | 82,704 | 10993.000 | 8916495.000 | 1059132.000 |
| allocator | plain | control | R2 | 30 | — | — | — | — | 82,736 | 10993.000 | 8916495.000 | 1059128.000 |
| allocator | media_rich | control | R1 | 30 | — | — | — | — | 803,328 | 13314.000 | 117608609.000 | 229300275.000 |
| allocator | media_rich | candidate | R1 | 30 | — | — | — | — | 803,452 | 13306.000 | 100831137.000 | 212522935.000 |
| allocator | media_rich | candidate | R2 | 30 | — | — | — | — | 803,596 | 13306.000 | 100831137.000 | 212522935.000 |
| allocator | media_rich | control | R2 | 30 | — | — | — | — | 803,328 | 13314.000 | 117608609.000 | 229300275.000 |

## Within-role normal repeat drift

| Corpus | Role | p50 | Mean | p95 | p99 | All fields within ceiling |
|---|---|---:|---:|---:|---:|---|
| media_rich | candidate | +0.059% | +0.060% | +0.077% | +0.056% | True |
| media_rich | control | +4.025% | +4.019% | +3.934% | +3.970% | True |
| plain | candidate | +0.179% | +0.154% | +0.164% | -0.607% | True |
| plain | control | +0.230% | +0.220% | +0.013% | +0.223% | True |

No normal repeat statistic exceeded its frozen drift ceiling; no repeat statistic is withheld.

## Normal diagnostic control/candidate deltas

| Corpus | Repeat | Metric | Delta | >5% trigger |
|---|---|---|---:|---|
| plain | R1 | p50 | +0.805% | False |
| plain | R1 | mean | +0.863% | False |
| plain | R1 | p95 | +0.689% | False |
| plain | R1 | p99 | +1.398% | False |
| plain | R1 | whole_process_rss_kib | -0.073% | False |
| plain | R2 | p50 | +0.755% | False |
| plain | R2 | mean | +0.797% | False |
| plain | R2 | p95 | +0.841% | False |
| plain | R2 | p99 | +0.559% | False |
| plain | R2 | whole_process_rss_kib | -0.010% | False |
| media_rich | R1 | p50 | -0.013% | False |
| media_rich | R1 | mean | -0.008% | False |
| media_rich | R1 | p95 | -0.024% | False |
| media_rich | R1 | p99 | +0.015% | False |
| media_rich | R1 | whole_process_rss_kib | -0.030% | False |
| media_rich | R2 | p50 | -3.825% | False |
| media_rich | R2 | mean | -3.814% | False |
| media_rich | R2 | p95 | -3.735% | False |
| media_rich | R2 | p99 | -3.750% | False |
| media_rich | R2 | whole_process_rss_kib | -0.191% | False |

## Allocator resource deltas

| Corpus | Repeat | Metric | Delta | >5% trigger |
|---|---|---|---:|---|
| plain | R1 | allocated_bytes | +0.000% | False |
| plain | R1 | allocation_calls | +0.000% | False |
| plain | R1 | region_peak_live_bytes | +0.000% | False |
| plain | R1 | live_bytes_before | +0.001% | False |
| plain | R1 | live_bytes_after | +0.001% | False |
| plain | R1 | rss_delta_bytes | undefined baseline | False |
| plain | R1 | whole_process_rss_kib | +0.107% | False |
| plain | R2 | allocated_bytes | +0.000% | False |
| plain | R2 | allocation_calls | +0.000% | False |
| plain | R2 | region_peak_live_bytes | +0.000% | False |
| plain | R2 | live_bytes_before | +0.001% | False |
| plain | R2 | live_bytes_after | +0.001% | False |
| plain | R2 | rss_delta_bytes | undefined baseline | False |
| plain | R2 | whole_process_rss_kib | -0.039% | False |
| media_rich | R1 | allocated_bytes | -14.266% | False |
| media_rich | R1 | allocation_calls | -0.060% | False |
| media_rich | R1 | region_peak_live_bytes | -7.317% | False |
| media_rich | R1 | live_bytes_before | +0.000% | False |
| media_rich | R1 | live_bytes_after | +0.000% | False |
| media_rich | R1 | rss_delta_bytes | -100.000% | False |
| media_rich | R1 | whole_process_rss_kib | +0.015% | False |
| media_rich | R2 | allocated_bytes | -14.266% | False |
| media_rich | R2 | allocation_calls | -0.060% | False |
| media_rich | R2 | region_peak_live_bytes | -7.317% | False |
| media_rich | R2 | live_bytes_before | +0.000% | False |
| media_rich | R2 | live_bytes_after | +0.000% | False |
| media_rich | R2 | rss_delta_bytes | +0.000% | False |
| media_rich | R2 | whole_process_rss_kib | +0.033% | False |

## Review triggers

- None above the frozen 5% diagnostic threshold.

Allocator elapsed comparisons are withheld. Whole-process RSS includes setup, verifier-visible process work, and teardown; operation allocator vectors include observer overhead. The 0424 experiment remains descriptive and does not establish a release latency or optimization claim.
