# Change 0418 performance evidence

The frozen verifier was replayed before this table was rendered. It validated
the retained raw reports, catalogs, corpus bindings, binaries, and the
published projections; a missing raw report may be represented by its zstd
fallback. Four process-isolated ABBA legs are shown separately for each
selector. These are within-process samples and matched ABBA repeat
observations; they make no host-population uncertainty or speedup claim.

`p50` is the harness integer-nanosecond midpoint (floored for an even
sample count); `p95` and `p99` are nearest-rank values. Latency cells are
milliseconds. A positive paired percentage means the candidate elapsed
value is lower than the matched control value under the verifier formula.

## Normal latency by ABBA leg

| Selector | Statistic | A1 control | B1 candidate | B2 candidate | A2 control | A1→B1 change | A2→B2 change | Control drift | Candidate drift |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| pptx_cross_copy_media_rich_lifecycle | p50 | 1175.449339 | 720.698255 | 723.741233 | 1172.463108 | +38.687% | +38.272% | -0.254% | +0.422% |
| pptx_cross_copy_media_rich_lifecycle | mean | 1175.287329 | 721.006278 | 723.114279 | 1172.576222 | +38.653% | +38.331% | -0.231% | +0.292% |
| pptx_cross_copy_media_rich_lifecycle | p95 | 1177.533392 | 724.745237 | 725.591962 | 1174.220584 | +38.452% | +38.207% | -0.281% | +0.117% |
| pptx_cross_copy_media_rich_lifecycle | p99 | 1178.774215 | 730.718957 | 726.061446 | 1175.103714 | +38.010% | +38.213% | -0.311% | -0.637% |
| pptx_cross_copy_plain_lifecycle | p50 | 9.925957 | 9.524911 | 9.545602 | 9.874076 | +4.040% | +3.327% | -0.523% | +0.217% |
| pptx_cross_copy_plain_lifecycle | mean | 9.936832 | 9.530486 | 9.559739 | 9.892925 | +4.089% | +3.368% | -0.442% | +0.307% |
| pptx_cross_copy_plain_lifecycle | p95 | 10.165784 | 9.759612 | 9.780996 | 10.043737 | +3.995% | +2.616% | -1.201% | +0.219% |
| pptx_cross_copy_plain_lifecycle | p99 | 10.252113 | 9.836602 | 9.949173 | 10.110843 | +4.053% | +1.599% | -1.378% | +1.144% |
| pptx_cross_copy_plain | p50 | 8.651347 | 8.178830 | 8.263475 | 8.578607 | +5.462% | +3.673% | -0.841% | +1.035% |
| pptx_cross_copy_plain | mean | 8.659232 | 8.190901 | 8.286710 | 8.590546 | +5.408% | +3.537% | -0.793% | +1.170% |
| pptx_cross_copy_plain | p95 | 8.839438 | 8.389026 | 8.566917 | 8.751308 | +5.095% | +2.107% | -0.997% | +2.121% |
| pptx_cross_copy_plain | p99 | 8.911298 | 8.487267 | 8.651807 | 8.793127 | +4.758% | +1.607% | -1.326% | +1.939% |
| pptx_cross_copy_media_rich | p50 | 1150.346612 | 683.943129 | 684.746665 | 1156.413381 | +40.545% | +40.787% | +0.527% | +0.117% |
| pptx_cross_copy_media_rich | mean | 1150.576694 | 683.979536 | 684.758025 | 1156.547504 | +40.553% | +40.793% | +0.519% | +0.114% |
| pptx_cross_copy_media_rich | p95 | 1152.754553 | 684.636500 | 685.229907 | 1158.359115 | +40.609% | +40.845% | +0.486% | +0.087% |
| pptx_cross_copy_media_rich | p99 | 1153.093815 | 685.032062 | 685.312009 | 1159.999753 | +40.592% | +40.921% | +0.599% | +0.041% |

## ABBA decisions

| Selector | Accepted statistics | Rejected statistics | Adverse in both pairs |
| --- | --- | --- | --- |
| pptx_cross_copy_media_rich_lifecycle | p50, mean, p95, p99 | none | none |
| pptx_cross_copy_plain_lifecycle | p50, mean, p95, p99 | none | none |
| pptx_cross_copy_plain | p50, mean, p95, p99 | none | none |
| pptx_cross_copy_media_rich | p50, mean, p95, p99 | none | none |

## Allocation and RSS guards

Allocator latency is excluded. Allocation fields are process-leg aggregate
observations from the allocator lane; phase selectors expose unavailable
allocation attribution. `allocated_bytes` and whole-process RSS pairs marked
REVIEW exceed the five percent review threshold. RSS is GNU `time -v` maximum
resident set size for the complete process, including setup and warmups.
Resource percentages use candidate-minus-control, so positive values mean
more candidate resource use.

| Selector | Pair | Allocation calls (control → candidate) | Allocated bytes (control → candidate) | Normal RSS | Allocator RSS |
| --- | --- | --- | --- | --- | --- |
| pptx_cross_copy_media_rich_lifecycle | a1_to_b1 | 1,998,631 → 1,852,353 (-7.319%) | 1,105,844,962,382 → 1,104,356,263,446 (-0.135%) | 819,028 → 884,576 KiB (+8.003%) REVIEW | 819,092 → 884,672 KiB (+8.006%) REVIEW |
| pptx_cross_copy_media_rich_lifecycle | a2_to_b2 | 1,998,628 → 1,852,346 (-7.319%) | 1,105,844,952,926 → 1,104,356,241,382 (-0.135%) | 820,056 → 886,172 KiB (+8.062%) REVIEW | 819,360 → 886,148 KiB (+8.151%) REVIEW |
| pptx_cross_copy_plain_lifecycle | a1_to_b1 | 1,598,970 → 1,490,430 (-6.788%) | 883,285,770 → 736,760,190 (-16.589%) | 82,740 → 82,720 KiB (-0.024%) | 82,672 → 82,736 KiB (+0.077%) |
| pptx_cross_copy_plain_lifecycle | a2_to_b2 | 1,598,970 → 1,490,430 (-6.788%) | 883,285,770 → 736,760,190 (-16.589%) | 82,736 → 82,620 KiB (-0.140%) | 82,704 → 82,736 KiB (+0.039%) |
| pptx_cross_copy_plain | a1_to_b1 | unavailable | unavailable | 82,608 → 82,656 KiB (+0.058%) | 82,676 → 82,556 KiB (-0.145%) |
| pptx_cross_copy_plain | a2_to_b2 | unavailable | unavailable | 82,732 → 82,616 KiB (-0.140%) | 82,732 → 82,668 KiB (-0.077%) |
| pptx_cross_copy_media_rich | a1_to_b1 | unavailable | unavailable | 819,344 → 884,676 KiB (+7.974%) REVIEW | 819,700 → 884,856 KiB (+7.949%) REVIEW |
| pptx_cross_copy_media_rich | a2_to_b2 | unavailable | unavailable | 820,596 → 886,132 KiB (+7.986%) REVIEW | 819,128 → 886,400 KiB (+8.213%) REVIEW |

The retained `live_bytes_before`, `live_bytes_after`,
`peak_live_bytes_before`, and `peak_live_bytes_after` values are literal
snapshots. They are audited as snapshots and are not interpreted as an
operation-local peak or a full output-retention memory measurement.

Deterministic within-process uncertainty details, including exact IID median
order-statistic intervals and seeded percentile bootstrap intervals, are in
[`uncertainty.json`](uncertainty.json). The full raw vectors remain in the
validated reports; [`summary.json`](summary.json) retains their checked statistics.

Verifier source: [`verify.py`](verify.py). Rendered from 0418.
