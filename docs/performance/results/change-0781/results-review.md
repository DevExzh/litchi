# 0781 paired-results disposition review

This is an independent read-only review of the completed 0781 packet at base
`fdca3e63037ac83ff27d4bf64a0f157b971b45f8`. It covers `plan.json`, the final
`analysis.json`, `observer-analysis.json`, and the archived candidate. The
captures, correctness gates, and source custody checks are complete. No
measurements were changed or resampled.

## Disposition

Reject the `Cow` candidate for production adoption and keep the restored base
as the live source. This agrees with the final disposition recorded in
`analysis.json` (`disposition.status = "rejected"`,
`production_change_retained = false`). The rejection is a performance-ROI
decision; the candidate passed the correctness and build gates.

The decisive result is the short-text many-shape case: 100 slides × 10 boxes
(1,000 boxes), as fixed by `probe-src/src/main.rs:96-101`. Its native p50
regresses in both end-to-end modes:

| Shape and mode | Median p50 change | Six-block p50 changes |
| --- | ---: | --- |
| many / write | **+12.121%** | +11.49% to +13.15% (6/6 regressions) |
| many / lifecycle | **+10.551%** | +9.33% to +14.04% (6/6 regressions) |

These values are in `analysis.json` under
`native.analysis.paired_by_block_before_after["many/write"].metrics.p50` and
`["many/lifecycle"].metrics.p50`. Their bootstrap ratio intervals remain above
one: `[1.115751, 1.129312]` for many/write and `[1.097564, 1.126413]` for
many/lifecycle. The same many-shape regression persists through mean, p95, and
p99, so it is not a tail-only or single-block anomaly. Many/write whole-process
RSS also rises 5.085% in the aggregate; RSS is a secondary gauge under the
packet's stated limits.

The native `regression_flags_over_5_percent` list has 20 entries, but its
coherent central and upper-quantile pattern is the many case. The remaining
rich/tiny tail and RSS flags vary by block or have intervals crossing one, so
they do not support a separate adoption decision.

The candidate's strongest benefit is the deliberately large payload fixture:
16 × 4 boxes with 40,000 ASCII units per box (`probe-src/src/main.rs:39,
96-101, 788-793`). Native p50 improves 25.555% for write and 19.186% for
lifecycle, with all six blocks improving. Unicode payload p50 also improves
8.314% and 8.073% in write and lifecycle. Tiny is effectively neutral at p50
(+0.748% and +0.881%), while rich is neutral and variable (+0.333% and
-0.476%); inconsistent rich/tiny p95/p99 flags do not establish a common
workflow effect. The complete native table is retained in
`analysis.json` under
`native.analysis.paired_by_block_before_after[*].metrics`.

The resource evidence explains the split without overturning it. Allocation
comparisons show no regression flags and no meaningful live/peak/net-live
change. Allocated bytes fall 19.443%/16.228% for payload write/lifecycle and
10.672%/9.628% for Unicode write/lifecycle, while many falls only
1.569%/0.946%. These are separate allocator-region measurements, not a reason
to ignore the elapsed regression. The Heaptrack diagnostic independently finds
the targeted conversion ancestry at 192 events and 7,680,000 requested bytes
before, versus zero after (`observer-analysis.json` at
`heaptrack.diagnostic.before/after.exact_attribution`). It is explicitly
diagnostic and carries no timing or causal claim.

The packet's controls are sufficient to make this disposition actionable. The
native lane has six alternating blocks, 30 samples, and three warmups; the
allocation lane has two blocks, three samples, and no warmup (`plan.json`). The
observer lane has 12 whole-process perf receipts with source/output parity, and
both Heaptrack traces parse and cross-check their totals. The final analysis
records 1,286 passed and zero failed tests across 34 suites, six quality gates,
and all verification flags true (`analysis.json` fields `test_summary`,
`quality`, and `verification`; `observer-analysis.json` fields `perf` and
`heaptrack`). The candidate's exact-output, refusal, and semantic checks pass,
so this is not a correctness rejection.

## Next bounded candidate

The next experiment should audit a narrower borrowed representation at the
plain-text conversion boundary, such as `Option<&str>` or a separate borrowed
publication view consumed directly by the ClientTextbox encoder. Keep the
existing owned paragraph path for non-left alignment, rich text, and PP9
smart-tag mutation. Avoid carrying a general `Cow` through every
`UserShapeData` and group helper unless the focused short-text matrix shows
that its overhead is gone. The candidate archive shows the current `Cow`
propagation in `candidate/applied/.../core/codec.rs:184-307` and
`.../escher/model.rs:426-473`.

Any follow-up must use the same ten-case matrix and alternating order, require
the many/write and many/lifecycle p50 results to stay within the existing
regression threshold, preserve exact output/refusal/semantic gates, and retain
the payload allocation win before considering adoption. The current packet's
payload improvement and Heaptrack result are useful motivation for that audit,
not an approval of this implementation.

The scope remains warm in-memory fresh PPT write/lifecycle. It makes no
cold-cache, device-floor, concurrency, CRUD-completeness, or general causal
claim; the observer counters include setup and verification, and RSS includes
the whole process (`analysis.json:limits`, `observer-analysis.json:limits`).
