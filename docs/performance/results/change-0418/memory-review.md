# Change 0418 memory tradeoff review

`review_scope: primary owned PPTX cross-copy lifecycle`

`recommendation: retain the latency path only with the existing ownership gate; do not make a memory-improvement claim`

## Recommendation

Retain the candidate reuse for the path already guarded by
`OpcPackage::is_unmodified_owned_source`: owned ingress, unmodified package
state, built-in part implementations, and default save options. Dirty
destinations, revoked authorization, caller-defined parts, and non-default
save options retain the clone-and-apply fallback. Existing typed provenance
refusals also remain in force. The
latency claim can remain scoped to the measured warm owned PPTX lifecycle.

The measured media result is large and consistent: lifecycle p50 falls from
1,175.449 / 1,172.463 ms in the controls to 720.698 / 723.741 ms in the
candidates, a 38.69% / 38.27% paired reduction. Mean, p95, and p99 move in the
same direction and the declared same-revision drift checks stay below 5%.
The plain lifecycle guard is positive in both pairings as well. That evidence
supports retaining this narrowly gated latency optimization, subject to the
memory limitation below.

## Observed memory tradeoff

The media workload crosses the protocol's 5% review threshold in both paired
normal runs:

| measurement | control | candidate | paired change |
| --- | ---: | ---: | ---: |
| normal whole-process RSS, A1 → B1 | 819,028 KiB | 884,576 KiB | +8.00% |
| normal whole-process RSS, A2 → B2 | 820,056 KiB | 886,172 KiB | +8.06% |
| allocator whole-process RSS, A1 → B1 | 819,092 KiB | 884,672 KiB | +8.01% |
| allocator whole-process RSS, A2 → B2 | 819,360 KiB | 886,148 KiB | +8.15% |

The allocator lane gives the same direction over 30 lifecycle samples. Mean
allocation calls fall from about 66,621 to 61,745 per operation, while
cumulative requested allocation volume changes only from about 36.8615 GB to
36.8119 GB (−0.135%). Requested allocation volume is an accounting total; it
does not describe simultaneously resident bytes or physical copies. Fewer
calls therefore do not offset the RSS result.

The allocator's literal end-of-region live snapshot rises from 222,238,733 to
287,885,862 bytes (+65,647,129, about 62.6 MiB). The before snapshot changes
from 171,451,284 to 169,821,461 bytes (−0.95%). These snapshots are global
process state at the end of the timed region, so they are useful paired
observations but are not an operation-local peak. The reported process
high-water snapshot is also not an operation peak: it was already sampled as
`peak_live_bytes_before` and remains unchanged after the operation. It rises
from 832,715,501 to 896,603,225 bytes (+7.67%), which is a process-lifetime
signal only.

The plain lifecycle does not show the same regression: whole-process RSS
changes are below 0.15%, allocation calls fall from 53,299 to 49,681, and
cumulative requested bytes fall from 29,442,859 to 24,558,673. This makes the
RSS increase shape-dependent and consistent with the media-rich candidate's
retained artifact, rather than evidence of a general memory improvement.

## Cause and bounds

`opened/cross_copy_plan.rs::build_candidate` first builds and serializes a
candidate, then reopens those bytes and returns the reopened `OpcPackage`.
`apply_plan` and forward `apply_patch` now assign that already validated
package. The prior path dropped the reopened package after planning and
reapplied the durable patch to a destination clone. Consequently, the
candidate path retains a complete serialized archive and a separately parsed
part graph (with sharing within the graph) as the assigned destination. The
operation timer and allocator region end after publication, before the local
package and sink values are dropped, so the end-of-region live increase is
expected to observe that ownership difference. The timer starts after corpus
cloning and sink reservation and includes package ingress, snapshots, planning,
application, and publication; RSS still covers the complete process and
preflight.

The path remains bounded by the existing checks: cross-copy preflight checks
the changed-part estimate, serialization uses the selected
`Limits::max_patch_bytes` (128 MiB by default), and OPC reopening applies the
existing package/read limits. Those are finite input, part, and output
limits, not an aggregate retained-memory budget. The current evidence does
not prove that a larger media package will refuse before resident memory
grows, and it does not establish an operation-local peak.

## Follow-up required before a broader memory claim

Keep the latency result and the observed RSS cost in the change record. A
future memory-focused batch should do both of the following:

1. Add an operation-local peak or retained-artifact measurement, and exercise
   near-limit and low-memory owned PPTX cases. The check should establish
   whether the existing finite limits are sufficient for the extra serialized
   archive plus parsed graph, or introduce a bounded refusal before assigning
   the candidate.
2. Evaluate a representation that can release the temporary serialized archive
   or decoded duplicate after validation while preserving exact-source
   preservation, physical fingerprints, patch validation, and the dirty/custom
   fallback. Any such change needs its own allocation and RSS evidence; the
   current cumulative allocation totals cannot establish it.

The coordinating review accepts retention of the proven owned-source fast
path: the substantial paired latency reduction and positive plain guards
justify the observed memory cost in this scoped batch. Authorization is
limited to latency; the media RSS increase is an accepted, workload-specific
tradeoff. It remains a regression for memory-sensitive callers and a priority
for follow-up measurement.

## Evidence

- [0418 change record](../../changes/0418-pptx-cross-copy-candidate-reuse.md)
- [allocation metrics](allocation-metrics.json)
- [source and compatibility review](source-review.md)
- [capture protocol](../change-0418/protocol.json)
- [published result bundle](README.md)
