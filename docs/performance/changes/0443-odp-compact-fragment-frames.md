# 0443: Compact ODP source-fragment scanner frames

The source-fragment scanner previously copied resolved namespaces and local
names into open-element frames, then used them only for fixed-name comparisons.
It now classifies each start/empty event while its namespace is in scope and
stores an exact element kind plus byte offset. Unknown markup remains opaque;
root bindings, retained source XML, spans, styles, limits and publication
readback retain their existing semantics. No public API or dependency changes.

The frozen A1/B1/B2/A2 matrix retains 24 reports/720 samples across
64/4,096/8,192 source slides. Medium/large allocation calls fall 14.886%/14.943%
in both repeats: 577,973 to 491,936 and 1,151,422 to 979,369. All 60 observations
per role/shape agree. Requested bytes fall only 2.557%/2.663%; peak and retained
live bytes are unchanged. The predeclared allocation-call gate passes.

The normal latency gate fails. Medium p50 changes -0.606%/+2.460%, large
+0.267%/+1.402% (R1/R2). The keep decision accepts those sub-5% normal p50 costs
for fewer allocation calls. Medium allocator R1 p95/p99 rise 13.310%/12.987%;
R2 falls 2.478%/2.414%, but the first tail observation remains unresolved.
All eight repeat flags remain, including a large baseline tiny timing shift.
No latency, peak, RSS or bounded-existing-append gain is claimed.

Four whole-process profiles show cycles +1.516% and instructions -0.739%,
including setup, warmups and oracle work. Both roles retain 13 symbolization
warnings per conversion. The original protocol inherited the prior batch's
freeze timestamp; its unchanged bytes and the actual capture chronology remain
in the [disclosure](../results/change-0443/freeze-review.json).
See [measurements](../results/change-0443/measurements.md),
[decision](../results/change-0443/decision.json) and the
[bundle](../results/change-0443/README.md).

The entire original scanner is retained byte-for-byte behind cfg(test).
Three differential tests compare 462 inputs, including every private span,
source byte, root binding, style name and exact error result. All 358 ODP tests
and strict owner Clippy pass. Full integration and portable validation receipts
are recorded in the bundle. No measured attempt is excluded or replaced.

Registry/default counts stay at 436/36, with no coverage promotion. Repeated
validation, one-shot lookups, bounded existing append, Part addition, repackaging,
native breadth, cold/range and scaling remain open under the full non-iWork goal.

Final gates pass: 358 ODP tests, 368 harness tests (one existing ignored),
strict owner Clippy, workspace/ODF feature checks, warning-denied rustdoc,
formatting and crate boundaries. Portable controls and 11/12 corruption probes
pass before/after cleanup. Four temporary executables totaling 1.83 GB were
removed; both build caches and user-owned GOAL.md were preserved.
