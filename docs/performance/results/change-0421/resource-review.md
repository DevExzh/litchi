# 0421 corrected allocator resource review

This bundle is a corrected current allocator baseline for the two named PPTX
lifecycle selectors. It is evidence that the post allocation live total is now
used for process high water accounting; it is not a control/candidate comparison
and makes no latency, speedup, memory reduction, or regression claim.

The four reports use the fixed source revision and binary, CPU 2, one worker,
fresh processes, 30 samples after three warmups, and the two declared repeats.
Corpus, catalog, output, source, binary, protocol, and verification identities
match within each selector's repeats. The summary's checks pass for peak-before
and peak-after being at least the corresponding live value and for chronological
peak monotonicity. These checks support the corrected process-counter
interpretation at the recorded quiescent boundaries.

## Compatibility and historical evidence

Every allocator report carries
`tool.allocator_counter_revision = "post_update_peak_v2"`; the normal binary
explicitly omits this allocator-only field. The capture, report-guard, summary,
normal-identity, and comparator checks reject a missing marker, an unknown
generation, or a markerless/new-generation mixture. Existing allocator policy
files do not pin this new marker, so they remain historical until both sides of
a future comparison are recaptured and the policy is intentionally versioned.

The old raw reports remain immutable. Their `peak_live_bytes_before/after`
values and any differences derived from them must not be compared with these
corrected observations. A historical value can happen to satisfy
`peak_live_bytes >= live_bytes` despite having been produced by the defective
pre-add update. The current R1/R2 rows are repeated observations of one fixed
revision, not a historical delta or an optimization result.

## Scope and remaining gap

The corrected fields are process-lifetime high-water snapshots of successful
allocator request sizes. They include setup and all allocator activity that
occurs before the boundary, exclude allocator-internal realloc copy overlap,
and do not measure RSS. `live_bytes_after` is the absolute live total at region
finish; subtracting `live_bytes_before` yields a signed net process change, not
bytes retained by this operation. The reports do not establish an
operation-local peak or an ownership-specific retained-memory budget.

The next safe measurement should add a separate region-local maximum initialized
from `live_bytes_before` and updated with each post-allocation/reallocation live
total. It must leave the process high-water counter untouched, define the
single-worker/background-thread synchronization contract, and retain explicit
drop boundaries for any claimed retained snapshot. Tests should cover a local
allocation below an older process high water, realloc growth followed by
shrink, deallocation of pre-existing bytes, and overflow/unavailable regions.
Only after that instrumentation is validated should a near-limit or
source-backed lifecycle use an operation-local peak in its resource analysis.

The full correction and compatibility contract is recorded in the
[0421 change record](../../changes/0421-allocator-peak-counter.md), with raw
identities and replay commands in the [bundle README](README.md).
