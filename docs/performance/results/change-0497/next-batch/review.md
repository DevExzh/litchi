# 0497 next-batch review: bounded source-backed multi-Part reads

This is a read-only review of `next-batch/design.md` and
`next-batch/implementation.patch` against accepted ADR 0005 and the current
OPC source. No production source, test helper, build, benchmark, or capture
was run or changed by this review. The protected spec-gap checkout was not
accessed.

The proposed seam is the right ownership boundary: an opt-in
`SourceBackedPackage::read_parts_ordered` can schedule calls to the existing
private `read_part` path while leaving ZIP validation, `PartCache`, source
freshness, and managed `PartData` in `litchi-opc`. The operation is not ready
to apply as written. The smallest safe next slice is a prepared-request
adapter with a local scoped worker team and an explicit scheduler workspace;
it must close the findings below before any performance claim.

## Existing contracts the adapter must preserve

The package owns the immutable source, indexed archive, read-ahead adapter,
catalog, and cache at
`crates/litchi-opc/src/source_backed.rs:5193-5213`. The one-Part entry point at
`:9741-9820` fences the source and execution context before cache admission,
keeps same-Part single-flight, and returns budgeted `PartData`. Its cold-load
path at `:9869-10083` fences again before exposing a decoded payload and rolls
back cache publication on a stale source or cancelled context. The batch layer
must continue to call that path; it must not call `IndexedArchive` directly or
construct an unreserved payload owner.

`read_source_at_with_context` at `:11134-11213` is the physical accounting
boundary. It reserves the requested `InputBytes`, shrinks the reservation when
the remaining budget is smaller, commits only accepted bytes (including short
reads), and charges interrupted retries as `Work`. The batch byte gate is a
logical declared-size scheduler limit. It must not precharge declared bytes as
`InputBytes` or replace this helper, otherwise overfetch and short-read
conservation will be wrong.

The current forward read-ahead adapter admits one forward owner at a time and
polls managed waiters at `read_ahead.rs:721-780`; its publication transition is
monotonic and drains the window at `:417-485`. A forward-policy batch can
overlap task/decompression work while physical `ReadAt` calls remain
serialized. Exact-policy and forward-policy measurements therefore need
separate labels. A batch must not add an operation-scoped mode guard or restore
forward mode after publication.

These rules follow ADR 0005's immutable positional `ReadAt`, stable source
version, hierarchical finite budgets, opt-in local CPU parallelism, and
measurement requirements
(`docs/adr/0005-io-memory-and-performance.md:8-61`).

## Blocking findings in the proposed patch

### 1. The public method does not compile

The source hunk in `implementation.patch:18-35` adds

```rust
self.read_parts_ordered_impl(partnames)
```

but the proposed `batch.rs` only defines the free function
`batch::read_parts_ordered` (`implementation.patch:218-221`). There is no
`read_parts_ordered_impl` in the current `SourceBackedPackage`. The smallest
seam is either to call `batch::read_parts_ordered(self, partnames)` directly
or to add the private forwarding method. This is a compile blocker before
behavioral review can be meaningful.

### 2. Preflight is not the one immutable request plan described by the design

The design promises input-order resolution and declared-size reads before any
worker. The patch's `preflight` returns only a sum
(`implementation.patch:276-310`); `fits_parallel` then resolves names and
metadata again, swallowing metadata errors with `.ok()`
(`implementation.patch:233-244`), and the parallel branch resolves and reads
metadata a third time (`:251-268`). This creates three problems:

* a central-directory error can be silently converted into a serial decision;
* source/context state can change between the checks and the actual request;
* the serial and parallel paths do not consume the same validated request
  identity, so their error boundary is not the promised deterministic one.

Make preflight return one bounded `PreparedBatch` containing, for each input
ordinal, the resolved catalog index (and entry id if useful) and declared
size, plus the checked scheduling sum. Use those records for both serial and
parallel execution. Do not use `.ok()` to decide policy. Admit the descriptor
storage before `try_reserve_exact`, and keep its reservation with the batch
workspace. A fence after the final metadata record and immediately before
dispatch makes the handoff explicit; each `read_part` call and the final
fence remain authoritative if the source changes afterward.

There are also early-fence gaps. `preflight` checks the request-count limit
before any source or context check, and a metadata error returns before its
post-read fence. An already-cancelled call can therefore return `Parts`, and a
source mutation racing a metadata failure can return the metadata error. The
public operation should check source then context before the count decision and
route every preflight/setup failure through the documented source/context
precedence. In particular, test source mutation during central metadata, not
only during decoded payload reads. An empty request should perform the agreed
source/context fences and return immediately without reserving control or
worker resources.

The patch also checks `ReadResource::TotalPartBytes` only when the `u64` sum
overflows (`:304-308`); it never checks the configured aggregate ceiling. If
the batch has an aggregate decoded-byte limit, call the checked limit with
defined duplicate semantics. If the existing limit is package-wide rather
than logical request-wide, do not silently reinterpret it: document a
separate batch request/byte ceiling. `ReadLimits::max_parts` is currently
documented as the admitted package-part ceiling, so using it as a request
count cap must likewise be an explicit API decision, including whether
duplicates count.

### 3. The stop flag does not stop workers waiting at the byte gate

`worker` fetches an ordinal and then calls `Gate::acquire`
(`implementation.patch:408-447`). `Gate::acquire` checks context, but not the
shared `stop` flag (`:167-197`). After one worker fails, another worker that
already fetched an ordinal can remain in the condition-variable wait and then
run the task after the failure. The extra work is still bounded by the worker
count, but it contradicts the stated stop-on-error contract and can turn a
fast deterministic failure into additional source/cache activity.

Pass the stop state into the gate or check it before and after acquisition;
notify gate waiters when stop is set. The permit must be dropped before a
worker returns. Use checked task increment/decrement invariants rather than
`saturating_add`/`saturating_sub` (`:182-204`), which hide a scheduler bug and
can make the observed gate counters appear valid. The worker wrapper should
catch the whole worker body, or otherwise prove that lock/index/allocation
panics cannot escape the scoped team. On a spawn failure, signal stop and
explicitly join every already-created handle before returning; relying on the
scope's implicit join does not make the error path's cleanup and precedence
observable.

The gate bounds active logical tasks and their declared sizes. It does not
bound retained cache payloads, read-ahead bytes, provider-side buffers, or
thread stacks. The tests and documentation must name this distinction.

### 4. Scheduler memory and object accounting is incomplete

The workspace reservation in `implementation.patch:74-134` is a useful
direction, but it is not yet a valid ADR 0005 memory proof:

* `try_reserve_exact` is used without checking the resulting vector capacity,
  although the design claims capacity checking (`design.md:101-108`). The
  existing read-ahead constructor checks capacity against its reservation; the
  batch vectors should do the same or use an allocation shape with a proved
  bound.
* `size_of::<T>().saturating_add(ALLOCATOR_ALLOWANCE)` can silently saturate
  before the checked multiplication. Use checked arithmetic throughout.
* `ERROR_BYTES = 256` does not bound the allocation of an `OpcError`; provider
  and preservation errors can carry dynamic strings or boxed causes. Either
  bound/charge the stored diagnostic, or retain a bounded error representation
  and reconstruct only an explicitly bounded public error.
* the request vector, worker-handle vector, mutex/condition state, atomics, and
  default OS thread stacks are not all covered by the stated reservation. Set
  a bounded worker stack policy and account for it, or state that stack memory
  is outside the resource guarantee and measure it in RSS. Do not claim a
  bounded scheduler resident footprint while leaving this implicit.
* the serial path still reserves result slots and `workers.max(1)` worker
  state, even though it allocates neither slots nor worker threads
  (`implementation.patch:245-250`, `:331-345`). This can refuse a one-Part
  serial read because of unused `Objects`/`Memory`. Reserve only the shape
  actually used; an empty request should reserve nothing.

The scheduler token must be acquired before every fallible allocation it owns,
and released after all handles, slots, permits, and temporary errors are gone.
The read-ahead window and cache payload reservations are separate existing
owners and must be included in budget snapshots without being charged again by
the batch workspace.

### 5. Failure cleanup does not mean cache rollback

The design correctly avoids rolling back cumulative `InputBytes` and `Work`,
but its cleanup language and proposed test oracle need a sharper boundary.
An earlier successful worker may publish a payload into `PartCache` before a
different ordinal fails. Dropping the result slot releases that `PartData`
handle; it does not necessarily evict the successful cache entry. The
immediate failure invariant can therefore be:

* no active gate permits, scheduler reservations, or cache flights;
* no returned partial vector or temporary slot ownership; and
* retained cache `Memory`/`Objects` may increase according to normal cache
  policy, then release on eviction or package drop.

An assertion that all retained bytes return to the pre-call baseline immediately
would require operation-owned cache entries and rollback/eviction, which is a
larger cache contract and should not be smuggled into this scheduler. Tests
must record the retained-cache delta separately from temporary-resource
cleanup. `InputBytes` and `Work` remain monotonic physical evidence; compare
their deltas and source-accepted bytes rather than demanding serial-counter
equality after a failure.

### 6. Final error precedence needs an explicit contract

`finish` checks source, then context, then returns the worker result
(`implementation.patch:449-456`). This gives source changes precedence, but a
cancellation arriving after a lower-ordinal runtime error replaces that error
with `Cancelled`; the design's “lowest ordinal runtime error wins” statement
does not say whether that is intended. Define and test one order for:

1. source mutation;
2. cancellation observed during or after the work; and
3. the lowest input-ordinal typed task error.

The same rule must cover setup errors (workspace reservation, allocation,
thread creation) and preflight errors. Existing `read_part` post-read ordering
is the useful model: source freshness must be checked before exposing a lower
archive or cache error. A worker panic should have a dedicated typed batch
failure or an explicitly accepted refusal; mapping it to
`SourceBackedOverlayUnavailable` conflates scheduler failure with the
publication/overlay domain.

## Concrete production shape for the next batch

Keep the public API on `SourceBackedPackage`; it already owns the source and
cache and already exposes an optional execution context at
`source_backed.rs:6046-6050`. The method should be an explicit opt-in and
should use that existing context, with no runtime or global pool parameter.

Use this sequence:

1. Fence source then context. Apply a documented request-count ceiling. Return
   an empty vector after the final fences without allocating scheduler state.
2. Acquire a bounded descriptor reservation and build one input-order
   `PreparedBatch` with `index`/`entry_id` and declared size. Check each
   metadata result, per-Part limit, and any defined aggregate limit in input
   order. Fence source/context after preparation.
3. Select serial versus parallel from the prepared sizes. Serial consumes the
   same prepared indexes and the existing `read_part`; it does not reserve
   worker slots. Parallel reserves only the slots/output/handles/control shape
   it will allocate, then starts at most
   `min(workers, max_in_flight_tasks, request_count)` scoped workers.
4. Give each worker an ordinal, stop flag, and RAII gate permit. Check stop
   before cursor claim, while waiting, and after admission. The gate counts
   declared bytes only; `read_part` remains the sole owner of payload
   `Memory`/`Objects`, `Work`, and physical `InputBytes` charging.
5. Join every started worker on all paths. Select the lowest ordinal among
   observed task errors only after the source/context precedence decision. Drop
   slots and scheduler reservations before returning an error. Return an
   ordered vector of managed `PartData` handles on success.
6. Keep publication separate. If a concurrent or later publication calls the
   existing monotonic `disable_read_ahead_for_publication`, the batch must
   tolerate forward-to-exact transition and must never restore the discarded
   window. Add a race test for this boundary rather than adding a temporary
   mode guard.

This keeps the proposed optimization small: it schedules the already-proven
one-Part operation and adds only bounded orchestration. It does not weaken
source freshness, cache ownership, or exact publication semantics.

## Focused acceptance tests

The proposed test file contains comments only, not executable tests. Before
benchmark hookup, add focused tests at the OPC owner/integration boundary for:

* compile/API forwarding, empty calls, unmanaged calls, one-worker calls,
  below-threshold calls, stable order, duplicate names, and exact semantic
  bytes;
* one preflight plan: missing name, metadata failure, count/byte limit, and
  cancellation are selected in input order without a hidden serial fallback;
* source mutation during metadata, cache hit, cold read, read-ahead fill, and
  final assembly, including a cached hit while `monitor_reads` is false;
* delayed independent Parts with exact policy proving task overlap, and
  forward policy proving the expected read-ahead physical serialization while
  recording logical-task overlap separately;
* task and declared-byte maxima, an oversized Part's serial fallback, a
  waiting worker stopping promptly after another ordinal fails, and cancellation
  while waiting at the gate;
* the lower input ordinal winning when higher ordinal work completes or fails
  first, with the selected source/context/cancellation precedence;
* scheduler `Memory`/`Objects` and cache-flight gauges on success, cancellation,
  allocation failure, thread-spawn failure, provider error, and panic; test
  retained successful cache entries separately from temporary cleanup;
* `InputBytes` deltas equal provider-accepted physical bytes for exact,
  short-read, and read-ahead/overfetch providers, with `Work` charged only by
  the existing cold-loader/retry path; and
* a concurrent monotonic publication transition, plus a static check that no
  global runtime or detached worker is created.

The production benchmark should follow these gates. First prove that the
timed route invokes `SourceBackedPackage::read_parts_ordered`, not the eager
`OpenSession` harness. Then collect workers 1/2/4/8 (capped by policy) for
few-large and many-small shapes, separately for exact and forward policies,
with source accepted bytes, read-ahead diagnostics, cache hit/cold/waiter
counters, gate maxima, retained budget gauges, allocations, RSS, and latency.
Only the exact-policy delayed source can support a physical-read overlap
claim; forward-policy results must account for the existing one-owner
read-ahead behavior. No scaling or Amdahl claim is justified until the
serial-equivalence, source-fence, error, and resource receipts pass.

## Review disposition

The design is directionally compatible with ADR 0005 and the current OPC
ownership model. The implementation patch is rejected for the next gate until
the missing forwarding seam, single prepared preflight, early/source-first
error paths, stop-aware gate, and complete allocation/accounting shape are
fixed and exercised. No production edit is requested in this review.

