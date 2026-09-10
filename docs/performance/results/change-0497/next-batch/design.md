# Draft design: bounded source-backed multi-Part reads

Status: design and review patch only. This directory is a proposal for the
next implementation after change 0497. It has not changed the production
workspace, has no acceptance result, and has no benchmark or scaling result.
The review patch is intentionally unapplied.

## Scope and overlap check

The proposed opt-in operation is:

```rust
impl SourceBackedPackage {
    pub fn read_parts_ordered(&self, names: &[PackURI]) -> Result<Vec<PartData>>;
}
```

The returned vector has one `PartData` for each input name, in the same order.
Repeated names remain repeated in the result. Equal repeated requests may share
the existing `PartCache` single-flight and cached payload. The method uses the
`ExecutionContext` already owned by the package and takes no runtime, pool, or
worker parameter.

The repository has no equivalent source-backed operation. `OpenSession::read_many`
is an eager `OpcPackage` adapter and returns raw ZIP results through
`soapberry_zip`; it does not enter `SourceBackedPackage`'s `PartCache`, source
version fences, or managed `PartData` ownership. `PartView::data` is currently
one-Part, and `PartCache` coordinates only same-Part cold loads. This proposal
therefore fills an open package-level seam rather than duplicating an existing
owner.

The method is deliberately opt-in. Existing constructors, `PartView::data`,
publication paths, and ordinary format APIs retain their current behavior.

## Contract and request plan

The operation has one source snapshot and one execution boundary for its whole
call. It uses this sequence:

1. Fence source, then execution context before making a policy decision. An
   empty request performs the same final fences and returns an empty vector
   without scheduler reservations.
2. Apply `ReadLimits::max_parts()` as the request-count ceiling. Every input
   occurrence counts, including duplicates. A count failure is sent through
   the final source/context precedence check.
3. For a managed package, reserve the bounded descriptor shape before growing
   the descriptor vector. Resolve every name in input order into one immutable
   `Request { index, declared_bytes }` record. Each metadata result is fenced
   for source freshness and context before it is exposed. Check each
   `PartBytes` ceiling and the checked aggregate `TotalPartBytes` sum after
   every record. The aggregate is a logical request sum, so a duplicate is
   counted twice even when the cache later shares one payload allocation.
4. Fence source then context after the final descriptor and immediately before
   dispatch. Serial and parallel paths consume those same prepared indexes;
   neither path resolves names or metadata a second time. A changed source,
   cancellation, missing name, metadata error, or limit error therefore has a
   deterministic preflight boundary.
5. Use the serial path when no context is present, there is one request, the
   effective worker count is one, the checked declared sum is below
   `min_parallel_bytes`, or any request exceeds `max_in_flight_bytes`. The
   serial path reserves only its output-handle shape and calls the existing
   `read_part` operation for each prepared index.
6. Otherwise run a local pull-based worker team. The team has at most
   `min(workers, max_in_flight_tasks, request_count)` scoped OS threads. An
   atomic cursor is the only queue. A worker checks the shared stop flag before
   claiming an ordinal, then acquires a permit for one task and its declared
   bytes before calling `read_part`. Gate waiters use a ten-millisecond
   condition-variable timeout and check the execution context on every wake.
   A worker error or cancellation sets the stop flag and notifies all waiters;
   workers already inside `read_part` are joined, while workers still at the
   gate do not start another task.
7. Join every started worker on spawn, cancellation, provider, allocation, and
   panic paths. Completed results remain in ordinal slots. The lowest ordinal
   among observed task failures is selected after the join. Source freshness
   is checked first, context cancellation second, and the selected setup/task
   error last. No partial vector is returned.
8. On success, perform the final source and context fences, then return the
   ordered managed handles. Temporary descriptors, slots, permits, worker
   handles, and scheduler reservations are dropped before the public return.

The existing `read_part` path remains authoritative for ZIP validation, cache
admission, decoded-size checks, cumulative `Work` and physical `InputBytes`,
cache flights, and `PartData` ownership. The declared-byte gate is only a
scheduler admission counter; it is not an additional payload reservation.

### Error representation and precedence

Preflight errors retain their existing `OpcError` variants and are selected in
input order. The parallel error slot has a fixed-size classification rather
than retaining an arbitrary provider or ZIP diagnostic in scheduler state. It
preserves `SourceChanged`, `Cancelled`, `ReadLimit`, resource-limit fields,
and I/O kind; dynamic provider/ZIP details are reconstructed as bounded,
content-free typed errors. This keeps the scheduler reservation proof finite.
A serial fallback returns the existing direct `read_part` error because it has
no cross-thread error slot. If exact dynamic diagnostics must be identical
between serial and parallel paths, that is an acceptance gate for a later
operation-local error transport; this proposal does not claim that property.

The precedence contract is explicit: source mutation wins over context
cancellation; cancellation wins over a setup or task error; otherwise the
lowest input ordinal among observed task errors wins. A worker panic is caught
inside the worker boundary and becomes the existing bounded source-backed
refusal, an explicit adapter refusal rather than a provider or publication
error. A provider that ignores cooperative cancellation can delay the join
until its existing `ReadAt` call returns; no forced interruption is claimed.

## Resource accounting

The scheduler and cache have separate ownership of their resources:

| Resource | Owner in this operation | Release/commit rule |
| --- | --- | --- |
| `InputBytes` | `read_source_at_with_context` through the existing `ReadAt` path | Charge accepted physical bytes, including short reads and overfetch; cumulative. |
| `Work` | `PartCache` for an elected cold loader | Charge declared decoded work once; same-Part waiters add no work; cumulative. |
| payload `Memory` | `PartCache` reservations carried by cache entries, flights, and `PartData` | Existing eviction and handle-drop rules remain authoritative. |
| payload `Objects` | `PartCache` payload/flight reservations | Existing flight completion and handle-drop rules remain authoritative. |
| descriptor `Memory`/`Objects` | Prepared request vector and its bounded owner | Reserve before the vector allocation; drop after the operation, including preflight errors. |
| scheduler `Memory`/`Objects` | Temporary output, slot, handle, gate, error, and worker-stack shape | Reserve before each owned allocation/spawn shape; drop after all workers and slots are gone. |

The descriptor reservation is acquired before `try_reserve_exact`. The
scheduler workspace uses checked arithmetic for every element-size,
capacity, worker-count, byte, and `u64` conversion. It includes allocator
allowance per vector element, fixed gate/error/control state, the bounded reconstruction text for stored
errors, a fixed scoped worker stack size, and the worker-handle vector. Every `try_reserve_exact`
capacity is measured against its charged envelope after allocation. Serial
work reserves only its descriptor and output shapes; it does not reserve
parallel slots, handles, or worker stacks. An empty request reserves nothing.

The task gate bounds active logical tasks and their declared sizes. It does
not claim to bound cache-retained payloads, read-ahead bytes, provider-owned
buffers, or memory outside the explicitly sized worker stacks. Those existing
owners remain separate and are measured independently.

The scheduler never reserves declared payload bytes a second time. There is no
TLS cache, detached worker, global pool, or unbounded result queue. All task
and byte counters use checked increments/decrements; a counter invariant is a
scheduler failure rather than a saturating repair.

## Freshness, cache, and retention

Every task calls the existing `read_part` path, including cache hits and
single-flight waiters. The batch layer does not call `IndexedArchive`
directly, bypass read-ahead, construct a bare payload owner, or restore a
forward read-ahead mode after publication. Publication remains a separate
monotonic exact transition owned by the existing helper; a concurrent
transition may drain the window while this operation is running.

A failed or cancelled call guarantees that no scheduler reservation, gate
permit, temporary slot, or cache flight remains active when it returns. A
successful worker may already have published a payload into the ordinary
`PartCache` before another ordinal fails. Dropping the batch result slot does
not evict that clean cache entry. Therefore retained cache `Memory`/`Objects`
may increase according to normal cache policy and release on eviction or
package drop. The cleanup oracle records this retained-cache delta separately
from temporary-resource cleanup. Cumulative `InputBytes` and `Work` are
physical evidence and remain monotonic; they are never rolled back by this
adapter.

## Proposed focused tests

The review patch contains a test-file placeholder only. The follow-up owner
must add executable tests beside the source-backed owner before any benchmark
hookup:

* API forwarding, empty calls, unmanaged calls, one-worker calls,
  below-threshold calls, stable order, duplicate names, and exact semantic
  bytes;
* one immutable preflight plan covering missing names, metadata failure,
  count/aggregate-byte limits, cancellation, and source mutation during
  central-directory metadata;
* source mutation during cache hit, cold read, read-ahead fill, and final
  assembly, including a cached hit with `monitor_reads` disabled;
* delayed independent Parts proving task overlap while exact-policy physical
  reads and forward-policy read-ahead serialization are recorded separately;
* task and declared-byte maxima, an oversized Part's serial fallback, and a
  gate waiter stopping promptly after another ordinal fails;
* cancellation before dispatch, while waiting at the byte gate, and during a
  delayed provider, with all workers joined and cache flights at zero;
* lower-ordinal failure selection after higher-ordinal completion/failure,
  source/context/task precedence, and bounded panic/spawn failure paths;
* scheduler `Memory`/`Objects`, cache-flight gauges, retained-cache deltas,
  and temporary cleanup after success, cancellation, allocation failure,
  thread-spawn failure, provider error, and panic;
* `InputBytes` deltas equal provider-accepted physical bytes for exact,
  short-read, and read-ahead/overfetch providers, with `Work` charged only by
  the existing cold-loader/retry path; and
* `OwnedSource`, `FileSource` (warm fixture), short-read, and delayed
  instrumented providers, plus a static assertion that no global runtime or
  detached worker is created.

The test oracle compares semantic bytes, deterministic error precedence,
cache-retention deltas, and temporary resource gauges. It does not demand
serial equality for cumulative physical counters or for retained clean cache
entries after a partial failure.

The phrase “budget counters at the same post-call state” in the 0497 next
implementation brief needs an explicit acceptance decision. Existing
`ExecutionContext` counters for `InputBytes` and `Work` are cumulative and
have no rollback API. Literal rollback equality with a serial failure would
require an operation-local child context threaded through `read_part`, which
is outside this scheduler slice. This draft keeps the smaller contract above;
if literal rollback of every cumulative dimension is required, reject this
proposal and specify that child-context design first.

## Measurement and acceptance gates

No baseline, scaling claim, Amdahl estimate, or production support claim is
made by this draft. The next batch must first pass compiler/clippy checks,
serial-equivalence, resource, cancellation, source-freshness, retention, and
deterministic-error tests. Only then should a benchmark owner add workers
`1, 2, 4, 8` within policy using the selected-Part serial loop as the control,
with few-large and many-small shapes, warm `FileSource`, instrumented
short-read, and fixed-delay providers. Exact-policy physical reads and
forward-policy read-ahead behavior must remain separate measurement labels.

## ADR verification

Before writing this draft, the accepted ADR set listed by
`docs/performance/results/change-0497/adr-refresh.json` was read. The
30 recorded SHA-256 entries match the current `docs/adr` files. The design
relies most directly on:

* ADR 0001 for explicit opt-in low-level APIs and typed failures;
* ADR 0003 for immutable `Send + Sync` snapshots, source identity, and
  concurrency boundaries;
* ADR 0005 for positional `ReadAt`, hierarchical budgets, bounded lazy state,
  and measured performance scope;
* ADR 0006 for fail-closed validation and deterministic errors;
* ADR 0008 for dependency-ordered gates and no unsupported performance claim;
* ADR 0010 for keeping archive implementation below facades;
* ADR 0011 for `litchi-opc` ownership of physical OPC reads; and
* ADR 0024 for the current `litchi-opc` topology.

No protected worktree or spec-gap checkout was accessed.
