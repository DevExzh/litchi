# Change 0498 production review

This is a source-only review of the frozen source-backed ordered multi-Part
read implementation. I did not edit production code or run a build, test,
lint, or benchmark command.

## Findings

The sole panic blocker from the first review is cleared. The parallel branch
now extracts the context with a fallible `ok_or_else` and returns
`SourceBackedBatchInvariant` if its scheduling invariant is ever violated;
there is no deliberate `expect` in the batch implementation.

The remaining worker boundary is typed: `Builder::spawn_scoped` returns an
I/O error, every admitted handle is joined, and a worker panic becomes
`SourceBackedBatchWorkerPanic { ordinal }`. The two new error variants are
included in the `OpcError` conversion match, and ordinary `ResourceLimit`
errors retain their original dynamic scope.

The wave contract is now coherent. A fence runs before every wave, all
admitted handles are joined, errors are reduced by input ordinal, and no later
wave starts after a member error. The final source then context fence runs
before either a successful `PartBatch` or an error is exposed. The prepared
path reuses managed entry metadata; the unmanaged path continues through the
ordinary serial `read_part` path.

`PartBatch` retains its structural `Vec<PartData>` memory and object
reservations until drop, while each `PartData` retains its independent cache
payload reservation. It exposes borrowing accessors and no unbudgeted managed
`Vec` escape. Scheduler control memory includes the per-worker two MiB stack
reservation, handle storage, and fixed control allowance. The two MiB value is
a conservative admitted stack budget for this implementation, not a universal
proof for every platform or future decoder; the deflated multi-Part stress
test now exercises the production worker path and should remain a gate if the
stack or ZIP backend changes.

The documented resource behavior is honest: `Work` and `InputBytes` are
cumulative actual charges and are never rolled back to imitate serial work;
cache hits and concurrent loader races can therefore change counters. Clean
payloads from a failed wave may remain in the ordinary cache. Unmanaged calls
have no hierarchical whole-collection reservation, retain serial one-Part
semantics, and are bounded by the new request-occurrence `max_parts` cap plus
the existing per-Part/read limits. Managed callers that need a tighter bound
on the retained collection must supply an execution budget.

The earlier occurrence-level `max_total_part_bytes` mistake is gone: batch
declared totals are used for parallel threshold selection only, while package
total limits remain catalog limits. Checked scheduler overflows use the batch
invariant error rather than fabricating a package read-limit failure.

## Disposition

No source-level blocker remains in the reviewed batch contract. Coordinator-
owned compilation, focused tests, formatting, lint, and documentation gates
remain necessary.
