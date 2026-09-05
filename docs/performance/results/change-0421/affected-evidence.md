# Change 0421: allocator peak-counter correction notice

Date: 2026-09-05

`performance_claim: none`

`claim_authorized: false`

This is an evidence correction notice. It does not rewrite retained raw
reports or authorize a memory claim.

## Finding and provenance

In [`allocation_metrics.rs`](../../../../tools/perf-baseline/src/allocation_metrics.rs#L334),
`Counters::live_add` passes the result of `checked_add` to `update_peak`.
`checked_add` calls `AtomicU64::fetch_update`; a successful `fetch_update`
returns the value observed before the update, while the atomic counter itself
has already been set to the post-add value. Consequently, `update_peak` sees
the pre-add live total. Both `peak_live_bytes_before` and
`peak_live_bytes_after` can therefore under-report the process high-water
counter. A paired difference can happen to cancel this error, but it cannot be
accepted without a corrected capture.

`git blame` identifies `4c0224bc44fb9dae881aa1a6bad696f62239373d`
(`perf(tools): add isolated allocator operation metrics`, 2026-08-20) as the
source-introducing revision for `live_add`, `update_peak`, and `checked_add`
([`checked_add`](../../../../tools/perf-baseline/src/allocation_metrics.rs#L406)).
`6f1f30e23e` later changed ordering and overflow handling but did not correct
the old-value return. The correction belongs in the allocator instrumentation,
with a regression test that observes a post-add high-water update.

For a measured, non-overflowing sample, this specific defect does not change
the successful `live_bytes` atomic update. Allocation/deallocation/reallocation
call and byte totals, `live_bytes_before/after`, and external process metrics
remain usable on their declared boundaries. `/usr/bin/time` RSS,
`peak_rss_bytes`, and Heaptrack's `peak_heap_bytes` are separate measurements;
they are not substitutes for a corrected `peak_live_bytes` vector.

## Claim-registry audit

The current [`claim-registry-v1.json`](../../claim-registry-v1.json) contains no
`peak_live_bytes` metric. Its only strict resource guardrail is
`claim-0251-xlsx-xml-borrowed`: the required metrics are Heaptrack allocation
and peak-heap/RSS values plus `/usr/bin/time` maximum RSS, all independent of
this counter. The other registry entries are latency-only, held, or rejected;
their retained claims do not depend on `peak_live_bytes`.

The accepted scoped latency records for 0400, 0401, 0410, 0413, and 0418 also
explicitly exclude allocator peak/live values from their authorization. Their
normal elapsed claims remain in scope. The 0418 registry claim is latency-only;
its resource and memory review is diagnostic.

## Evidence in correction scope

Examples of retained bundles that contain the affected allocator field directly,
or derive a projection from it: `0271`, `0395`, `0397`, `0400`, `0401`, `0402`,
`0406`, `0408`, `0409`, `0411`, `0412`, `0417`, `0418`, `0419`, and `0420`.
This list is not exhaustive: all pre-repair reports emitted by the affected
allocator implementation are in scope. The 0410 compressed allocator capture and its change record also report a
lifetime peak-after value. These are evidence-level impacts, not automatic
claim revocations.

The practical disposition is:

| Evidence | Peak-counter disposition | Values that remain usable now |
| --- | --- | --- |
| 0417 representative baseline and its two allocator selectors | Re-capture peak fields before treating them as high-water evidence | Normal timing, allocator calls/bytes, and external RSS |
| 0418 lifecycle package and memory review | Keep the latency claim; withhold every peak-live number and derived peak comparison pending corrected ABBA allocator captures | Normal timing, `live_bytes_after` as a live snapshot, and GNU-time RSS |
| 0419 resource diagnostic | Do not use its peak-live equality/delta as evidence; no claim was registered | Requested allocation totals, calls, live-after snapshots, and RSS on their declared process scope |
| 0420 resource diagnostic | Withdraw the reported media high-water reduction and plain high-water delta pending corrected captures; no latency claim is affected | The live-after and RSS observations remain usable for their stated process scope |
| Earlier direct-key bundles listed above | Preserve raw files, but mark peak-live projections as requiring corrected regeneration if they are reused | Their non-peak counters and independent metrics, subject to each record's existing boundary |

No retained claim should be reconstructed by subtracting a repaired value from
an old value. Re-run both control and candidate allocator roles with the fixed
instrumentation, then regenerate the affected projections and review any
peak-derived decision. The existing raw reports remain historical artifacts.

## Policy compatibility requirement

`perf-regression-policy-allocator-v1.json` and
`perf-regression-policy-xlsx-allocator-v1.json` require both peak-live paths;
the default policy lists them as optional but still recognizes the paths. The
report identity now includes
`tool.allocator_counter_revision = "post_update_peak_v2"` for corrected
allocator binaries. The existing comparator checks the complete tool identity
against policy and between reports, rejecting an old/new mixture or either
generation with the wrong policy. Normal binaries omit the field. The added
compatibility test covers these cases without changing comparator production
code.

Legacy markerless policies can still replay historical pairs for schema
compatibility; that does not authorize reuse of defective peak fields. Both
sides must be freshly captured with corrected instrumentation before a policy
opts into the new identity and a peak guardrail is reviewed. Call, byte, live,
overflow, status and cardinality validation remain required.

This notice does not claim completion of the broader performance goal.
