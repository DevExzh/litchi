# 0732 empirical results review

The frozen packet passes this bounded empirical review. Its arithmetic and
interpretation match the retained plan, analyzer output, raw captures, and
same-owner hash calculation. The evidence supports native phase attribution
for the warm `45543.ppt` slide-removal workflow in the diagnostics-enabled
release binary. It does not support an optimization or speedup claim, exact
ordinary-build phase fractions, or a default-build code-generation claim.

## Matrix and controls

The plan is three cycles by three rounds by four routes, with 50 measured
samples and three warmups per process: 36 processes, 1,800 measured
lifecycles, and 108 warmups. The fixed comparison chain is ordinary opaque to
ordinary split, ordinary split to profiled empty, and profiled empty to
profiled clock. Route order rotates and reverses by round. All measured
outputs pass the sealed PPT oracle; the analyzer reports no failed lifecycle.

The process-level p50 range for ordinary opaque is **978.53–997.45 µs**.
Recomputing the same-round process p50 and mean ratios gives:

| Control pair | p50 change | Mean change | Interpretation flags |
| --- | ---: | ---: | ---: |
| Ordinary opaque → ordinary split | +0.566% to +3.370% | −0.020% to +0.839% | 0 |
| Ordinary split → profiled empty | −0.847% to +0.660% | −0.152% to +0.510% | 0 |
| Profiled empty → profiled clock | +4.077% to +5.382% | +2.094% to +2.337% | 3 p50 |

The three flags are exactly the three cycle-0 profiled-empty-to-clock
comparisons. All 27 pair comparisons and 54 central checks remain retained;
none was selectively rerun. The fixed recorder's separate calibration has a
490–520 ns process p50 and is reported without subtraction. That is much
smaller than the full route difference, while clock dispatch, instrumentation,
and compiler layout can affect more than the direct timestamp callback.

## Phase arithmetic

The 450 profiled-clock owners each contain a complete 20-event trace, giving
9,000 retained event records. Event spans stay within their commit windows.
The phase ranges in the main report are consistent with the raw samples:

* `EmbeddedFinish` is 188.49–190.83 µs and 16.73–16.87% of same-owner whole
  time.
* `PublicReopen` is 104.75–105.69 µs and 9.68–9.84% of same-owner whole time.
* `ArtifactHashBefore` is 16.51–16.66% and `ArtifactHashAfter` is
  16.73–16.89% of same-owner whole time.
* Combining the two hash spans per owner before taking medians gives
  350.06–350.49 µs, or 33.24669–33.87107% of whole time (reported as
  33.25–33.87%). The equivalent commit fraction is 42.77–43.40%.

The combined hash result therefore must not be obtained by adding the two
component medians. The retained `report-stats.json` calculation uses each
owner's two spans and then takes the process median, which is the appropriate
denominator for this comparison. The observed commit occupies 77.07–77.27%
of whole time, with 54.83–56.04 µs of median commit residual outside the ten
reported spans.

These are profiled-clock route observations. Because the clock control has
three p50 flags, they cannot be transferred as exact phase fractions to the
ordinary opaque route. The phase table is useful for ranking work in the
observed route; it is not evidence of savings available by deleting any span.

## Source and behavior boundary

The source custody record and ordinary-method check show that the ordinary
PPT commit is byte-identical. The feature-gated diagnostic copy retains the
existing publication order, source and working snapshots, live-record and
payload owners, patch and inverse behavior, both required artifact hashes,
public reopen, unrelated-stream validation, payload checks, and output oracle.
`StructuralNoOp` marks an empty document patch only; formatting-only work can
still change the complete artifact. The observer is synchronous and
content-free, and the probe records events externally in a fixed-capacity
stack. The evidence therefore attributes observed phase spans without making
the callback part of the document result.

The packet includes successful feature-off and feature-on quality gates,
qualification, 36 native processes, independent replay, and 23 corruption
controls. Failure outcomes for wrapped operations and the existing typed
error paths are covered by the source tests; explicit residual checks after a
read-only phase remain residual work rather than being mislabeled as a phase
failure.

## Bounded next investigation

The next bounded measurement should inspect the work inside `EmbeddedFinish`
on the same fixture and CPU, using opt-in sub-boundaries or an external
profile while retaining the complete finish operation. A separate bounded
hash investigation may examine the two native digest calls, but it must keep
both calls, their inline evaluation order, patch construction, validation,
public reopen, exact output oracle, source sharing, and reversible patch
checks. Any candidate should be compared against the ordinary public path and
the four-route controls before drawing a performance conclusion.

The observer control should remain explicit: report empty-observer versus
clock-observer deltas and calibration independently, and refine that control
before publishing exact ordinary phase fractions. No validation bypass,
digest removal, unbounded cache, save-policy change, or speedup claim follows
from this packet.

## Scope limits

This is one warm in-memory PPT fixture and one slide-removal operation on one
CPU in a virtualized host. It does not establish producer breadth, cold I/O,
allocation or RSS behavior, concurrency behavior, or a general speedup. The
ordinary routes and profiled routes run in one diagnostics-enabled binary, so
the controls include split, compiler, and observer effects and do not measure
default-feature code generation. Process p50 ranges describe repeat spread,
not confidence intervals; all raw samples and tails remain in the packet.

Evidence: [main report](../../0732-ppt-native-phase-attribution.md),
[analysis](analysis.json), [same-owner hash calculation](report-stats.json),
and [packet replay instructions](README.md).
