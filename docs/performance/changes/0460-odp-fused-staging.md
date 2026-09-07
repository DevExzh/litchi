# 0460 — retain fused ODP staging and source scanning

The private ODP staging path now feeds settings, declarations, page metadata and
source-fragment state from one namespace-aware event stream. It retains the
existing state machines, historical error precedence, BOM-relative source
spans, lexical preservation, source ownership, candidate readback, patch
identity and no-op behavior. The source read review found no semantic blocker.

The frozen A1/B1/B2/A2 lifecycle matrix retains 24 reports and 720 samples.
Normal candidate-minus-baseline p50 deltas are -3.5430% / -3.8257% for tiny,
-6.0849% / -5.0679% for medium, and -5.4732% / -5.8268% for large in R1/R2.
The predeclared keep gate applies to medium and large: all four p50 deltas
exceed the 3% improvement threshold and their independent 95% bootstrap upper
bounds remain below zero. No adverse >5% elapsed or process-RSS flags are
present; the largest process-lifetime RSS delta is +3.6357%.

Allocator p50 elapsed deltas are -4.7840% / -4.1309% for tiny,
-7.3100% / -6.7785% for medium, and -7.3197% / -6.9282% for large. Across all
six allocator lanes, allocated bytes fall by 4,642, allocation calls by 16,
reallocation calls by 12, and deallocation calls by 4. Regional peak above
entry and retained-live deltas are unchanged. These allocator observations are
operation-scoped and do not establish a whole-process memory bound.

The supplementary large-input phase clocks show transaction p50 reductions of
24.2681% / 24.0697% in R1/R2. Add, snapshot-open, commit and publication phase
changes are small by comparison; these clocks are separate mechanism evidence,
not the lifecycle keep decision. Separate whole-process counters report
instructions -6.1959%, cycles -4.9100% and branch misses +0.5503%; setup,
warmups and checks are included, so these are not operation-only or causal
counts.

The corrected owner retry passes 371 ODP tests and warning-denied all-target
Clippy passes. The optimization is retained under the scoped 0460 evidence.
All builds, 387 harness tests (one ignored), documentation, boundaries,
portable verification and owned temporary cleanup pass. This change adds no selector or corpus
coverage; the registry remains 439 selectors / 36 defaults, the full non-iWork
goal remains open, and iWork remains outside scope. See the [comparison
summary](../results/change-0460/summary.json), [phase summary](../results/change-0460/phase-summary.json),
and [source review](../results/change-0460/source-review.md).
