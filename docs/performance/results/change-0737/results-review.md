# 0737 lifecycle qualification results review

## Scope and evidence

The frozen matrix completed 138 serial processes: 108 native processes with
4,500 samples from the five 50-sample arms and 18 one-sample fresh controls,
plus 30 allocation processes. Both qualified PPT fixtures were used, with
CPU 12 pinning, rotated process order, and the predeclared paired process
units. The native lane contains 4,518 samples in total.

The analyzer and an independently implemented audit both passed and matched the
full process-level statistics. Every inspected raw report had the expected
source/output hashes, eight rejected corruption controls, `oracle_ok == true`
with empty failure reasons, and the required lifecycle receipt fields. The
offline contract rejected all 17 deliberate report corruptions. The capture
and audit receipts bind the frozen source, binaries, fixtures, manifest, and
preflight. No production candidate was measured, and the packet makes no
claim about the cause of the rejected 0735 candidate.

Percentages below are paired process changes, `100 * (right / left - 1)`, and
the table reports the median of nine pair changes. Bootstrap intervals use
the frozen seed 7337 and 10,000 resamples. A flag means the absolute change
exceeded 5%; pair indices are the predeclared repeat indices.

## Native full-window results

| Fixture and comparison | p50 median [bootstrap 95% interval]; pair range | Mean | p95 | p99 / maximum | Full-window flags |
| --- | ---: | ---: | ---: | ---: | --- |
| Primary archive → legacy-a | +0.724% [0.360%, 1.266%]; [0.277%, 1.360%] | +1.878% | −0.039% | +0.001% / +0.001% | none |
| Primary legacy-a → legacy-b | +0.109% [−0.117%, 0.217%]; [−0.580%, 0.245%] | +0.009% | +0.044% | −0.362% / −0.362% | none |
| Primary legacy-a → strict-retained | −14.000% [−14.261%, −13.529%]; [−14.507%, −13.504%] | −7.916% | −6.690% | −6.751% / −6.751% | p50, mean, p95: 9/9; p99/max: 7/9 (0–5, 7) |
| Primary strict-retained → strict-drained | +10.099% [9.769%, 10.433%]; [9.747%, 10.567%] | +1.902% | −2.103% | −0.042% / −0.042% | p50: 9/9 |
| Secondary archive → legacy-a | +6.668% [5.902%, 7.486%]; [5.756%, 7.643%] | +2.519% | +0.528% | −0.105% / −0.105% | p50: 9/9; p99/max: pair 6 |
| Secondary legacy-a → legacy-b | −0.125% [−0.420%, 0.281%]; [−0.525%, 0.284%] | −0.057% | −0.260% | +0.459% / +0.459% | p99/max: pair 6 |
| Secondary legacy-a → strict-retained | +0.325% [0.010%, 0.598%]; [−0.161%, 0.697%] | +0.795% | +1.267% | +1.533% / +1.533% | p99/max: pair 6 |
| Secondary strict-retained → strict-drained | −0.313% [−0.480%, 0.069%]; [−0.590%, 0.161%] | +0.159% | −0.681% | +0.063% / +0.063% | none |

The archive-to-new-legacy comparison is close on the primary fixture but
shows a repeatable **+6.668% secondary p50** change: all nine pairs exceed the
review threshold. This comparison spans the independently rebuilt original
probe and the revised controls binary, so it combines binary, observer, and
serialization differences. It is not evidence of a production regression or
of a lifecycle mechanism. The duplicate legacy A/A comparison is small in
p50 on both fixtures. Its one p99/maximum outlier on each secondary control
comparison remains recorded and does not change the p50 conclusion.

The primary strict-retained to strict-drained contrast is large and consistent
in the full window: drained is slower by 9.747–10.567% in all nine p50 pairs.
The corresponding secondary contrast is small and inconclusive, with a
−0.313% p50 and an interval crossing zero. This fixture dependence prevents a
general transfer of the primary result to an ordinary production path.

The legacy-a to strict-retained primary contrast is also large, but it changes
validated warmup behavior and per-sample receipt construction while both arms
retain full witnesses. It therefore does not isolate retention. The strict
retained-versus-drained pair is the narrower retention diagnostic because the
strict arms share the owner operation, warmup validation, full oracle cadence,
and receipt path.

## Ordered samples and fresh controls

The endpoint first-ten/last-ten contrast in strict-drained is near zero on
both fixtures: the primary median is **+0.100%** across a range of −0.209% to
+0.512%, and the secondary median is **−0.114%** across −0.998% to +0.640%.
That endpoint result does not establish a stationary sample distribution or
show that sample-order bands disappeared. The retained and drained arms still
show distinct ordered trajectories, and the process-level plot preserves the
repeatable within-process bands. The endpoint windows are descriptive; they are neither independent process
replicates nor a replacement for the original full-window comparison.

The fresh arm contributes 18 one-sample process controls. It is not a
50-sample distribution and has no tail comparison. Its `samples=1` setting
also changes the initial receipt/witness vector capacity relative to the
50-sample strict arms, so it combines fresh-child behavior with a different
initial buffer layout and process length. It is useful as a bounded control,
not as a pure fresh-child causal test.

## Allocation lane

All 30 allocation reports have identical owner-region fields within each
fixture, across archive, legacy, zero-warmup drained, three-warmup drained,
and three-warmup retained arms:

| Fixture | Allocated | Deallocated | Calls | Peak live | Retained |
| --- | ---: | ---: | ---: | ---: | ---: |
| Primary | 11,391,768 | 11,001,624 | 5,658 | 1,992,885 | 390,144 |
| Secondary | 10,444,631 | 10,159,447 | 18,662 | 1,976,223 | 285,184 |

Every allocation comparison has zero difference for every field. These are
boundary-relative allocator counters for the owner region; they are not heap
size, RSS, or an explanation of the native timing result. The allocation lane
supports no latency claim and supplies no allocation or cache cause for the
0735 secondary regression.

## Disposition

0737 characterizes the lifecycle harness and preserves the original 0735 rejection
without changing production. The primary fixture shows strong sensitivity to
full witness lifetime, while the secondary p50 contrast is inconclusive, and the archive
versus rebuilt-legacy difference exposes an observer/build/layout effect that
must be isolated first.

The next retry should preserve the legacy `Sample` layout and serialization
path while qualifying that observer effect against the archived binary. The
strict-drained arm remains a diagnostic lifecycle control; its primary result
must not be transferred to ordinary production behavior. No candidate
reinstatement or causal explanation of 0735 is supported by this packet.
