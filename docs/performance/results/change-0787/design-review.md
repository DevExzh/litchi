# 0787 cached Part scheduling design review

This is a read only design and source audit for the next candidate after the
finite-budget read baseline in change 0786. It does not implement a
production change, run Cargo, run a native workload, or run a profiler. The
review is based on the production tree at `6b47617020` (whose production
source is byte-identical to the measured 0786 revision).

## Finding

The measured Parts control is a real scheduling candidate. In the large,
zero-floor, primed case, the 0786 report measures a width-one p50 of 3.045 us
and a width-32 p50 of 578.022 us; the paired speedup at width 32 is 0.005244x.
The observer reports zero source calls for every primed Part sample. Thus the
timed work is cache-hit handling plus the parallel scheduler, and not source
I/O or decompression. The 0786 source review identifies the relevant
eligibility gate: declared member sizes admit the parallel branch before
`read_part_prepared` discovers the cache hit.

The smallest compatible candidate is a scheduling hint over the cache's
existing state:

1. Keep `PreparedBatch::build`, `OutputAdmission::reserve`, the selected
   `SchedulerAdmission`, the context cancellation check, and the cumulative
   `Resource::CpuTasks` charge exactly where they are.
2. Keep the existing `read_parallel` entry fence. For a package with no
   caller-supplied `ScopedWorkers`, inspect the current cache state under the
   existing cache mutex. If every prepared `EntryId` has a completed cache
   entry and no active or provisional flight, take a private serial execution
   path that still calls `read_part_prepared` for every request.
3. If any entry is absent, active, or provisional, use the current parallel
   path byte-for-byte in structure. If an entry changes after the hint, the
   serial path falls back naturally through the ordinary `read_part_prepared`
   cache admission and may perform a cold load.
4. Keep the existing scheduler reservation alive until the operation returns.
   The hint therefore does not narrow admission and does not release or skip
   `Workers`, `IoConcurrency`, scheduler `Memory`/`Objects`, output
   `Memory`/`Objects`, or `CpuTasks` accounting. It removes worker/channel
   construction only when the hint is true.

This must not be implemented as `parallel = false` in
`read_parts_ordered`. That would move the decision before
`SchedulerAdmission::reserve` and change typed refusal and reservation
behavior. It must also not call the current `read_serial` helper without
qualification: the private parallel routes catch worker unwinds and return
`SourceBackedBatchWorkerPanic { ordinal }`, whereas ordinary serial reads let
an unwind propagate. The cached private path needs a caught serial helper,
with the same wave boundaries and lowest-ordinal error selection as the
existing worker routes, or the optimization must be rejected.

## Existing seams and invariants

`read_parts_ordered` in `crates/litchi-opc/src/source_backed/batch.rs` first
captures the package context and calls `fence`, rejects an empty or oversized
request, builds immutable `PreparedRequest` records, reserves output slots,
and computes the aggregate/per-task eligibility gate. Once the gate is
true, it tries widths from the requested width down to two with
`SchedulerAdmission::reserve`. That admission reserves `Workers`,
`IoConcurrency`, scheduler `Memory`, and scheduler `Objects` before any
task starts. Only after it succeeds does the function check cancellation and
consume one `CpuTasks` unit per prepared request. These are the boundaries
the candidate must leave intact.

`read_parallel` then calls `fence` again. Without a caller facility it uses
`wave_end` to choose one wave or the reused-worker path. `read_one_wave`
translates a joined worker panic to the lowest input ordinal; the reused
worker loop catches each request panic, drains the admitted wave, and applies
the same ordinal selection. The caller-facility route wraps each task in
`catch_unwind`, stores results in input slots, and calls
`ScopedWorkers::run_all`.

`PartCache::enter_with_observer` is the authoritative cache transition. It
holds `PartCache::state` while it first checks `flights`, then checks
`entries`, advances the LRU clock, and either returns a hit, joins a flight,
or reserves a cold load. A cache entry can be provisionally present while
its publication flight is still active (`pending`), so an entry lookup alone
does not establish a completed hit. The hit path clones the existing
`CachedPayload`, then `read_part_with_observer_and_capture` checks source
freshness and execution context before returning `PartData`.

The proposed hint should therefore be a private, allocation-free predicate on
`PartCache`, conceptually:

```text
lock state
for each prepared entry_id:
    if flights contains entry_id: return false
    if pending contains entry_id: return false
    if entries does not contain entry_id: return false
return true
```

It must not advance the LRU clock, increment `hits`, clone a payload, create a
`PartData`, reserve a budget resource, or install a new map entry. The
`pending` check is conservative even though current publication invariants
normally pair it with a flight. The method should recover a poisoned mutex in
the same way as the existing cache methods, and it should remain private to
the cache/batch implementation; `cache_diagnostics` cannot serve as a query
because it exposes no member identities and is racy for this purpose.

There is no useful stable per-entry atomic to add. Cache entries can be
evicted, pinned by `PartData` or `into_arc`, superseded by a flight, or
invalidated after a source-version change. An atomic memo or a hit guard
would either be a second hidden cache state or would alter eviction/pinning
and reservation behavior. The current mutex-protected map is the authority;
the predicate is only a hint and the normal read path remains authoritative.

## Why the race is safe

The predicate does not pin anything. Another operation can evict an entry,
start a flight, or make a previously clean entry provisional after the lock is
released. That only makes the hint stale. Every request still calls
`read_part_prepared`, whose existing path performs the cache transition and
all source/context checks. A stale hint can therefore cause a serial cold
read or a flight wait, but cannot return stale bytes, skip a limit, or expose
a partial `PartBatch`.

The initial `read_parallel` fence must remain immediately before the hint.
Each hit still runs the ordinary post-entry `source.ensure_current` and
`cache.check_context` checks. A source revision or cancellation observed
after the hint therefore has the same final-fence precedence as the current
parallel path. The outer `finish` fence remains unchanged. A source change
or cancellation during a race must be tested with a source/token that changes
between the hint and the first hit; the expected result is the existing typed
`SourceChanged` or `Cancelled` error and no returned batch.

The scheduler is intentionally retained even for a true hint. This makes
the candidate conservative: a finite `Workers` or `IoConcurrency` refusal,
or a finite scheduler `Memory`/`Objects` refusal, remains a refusal at the
same admission point. The output reservation is also retained by
`PartBatch`, and the cumulative `CpuTasks` charge still occurs before any
request is read. A cache hit continues to consume no `Work` or source input
bytes, exactly as it does today.

## Caller-provided `ScopedWorkers`

The fast path should be restricted to
`context.scoped_workers().is_none()`.

With a caller facility, `SchedulerAdmission` already treats worker handle and
stack memory differently, and `read_with_scoped_workers` creates one caught
task per request and invokes the facility once per admitted wave. Bypassing
that call for all cache hits would be observable to a caller-owned executor,
would skip its task/panic boundary, and would provide no private thread or
channel creation benefit. The accepted execution contract says that the
facility runs borrowed tasks to completion; retaining the existing route is the
safest interpretation of that contract and preserves task count, ordinal
errors, cancellation timing, and user-supplied scheduling behavior.

For the private-worker route, the cached serial helper must retain the
parallel route's unwind translation. It should process the same `wave_end`
segments, catch each request closure, map a panic to
`SourceBackedBatchWorkerPanic { ordinal }`, drain the current logical wave,
select the lowest ordinal error, and stop before the next wave after an
error. A helper that simply calls `read_serial` is insufficient even though
normal cache hits do not call `ReadAt::read_at`: every hit still calls source
version checks, and a caller source is permitted to unwind unless the path
catches it.

## Refusal and preservation matrix

| Behavior | Candidate requirement |
| --- | --- |
| Empty request, request-count and declared-size preflight | Unchanged before the hint. |
| `Workers`/`IoConcurrency` zero or insufficient grants | Existing scheduler/serial admission runs first; no hint may bypass it. |
| Scheduler `Memory`/`Objects` refusal | Same width fallback and final typed refusal. |
| Output `Memory`/`Objects` refusal | Still occurs before cache inspection. |
| `CpuTasks` | Consume exactly the prepared request count before the hint, including all hits. |
| `Work`, `InputBytes`, ZIP read and decompression | Cache hits remain free; a race-induced cold read follows the current path. |
| Source revision | Every returned hit still performs the existing source fence; stale hints cannot publish. |
| Cancellation | Existing initial, per-request, per-wave, and final checks remain authoritative. |
| Cache eviction | Predicate never evicts or updates LRU; ordinary `enter` remains the only transition. |
| Cache pinning | Predicate clones nothing; returned `PartData` pins exactly as before. |
| Duplicate request names | The same entry may satisfy every occurrence; output order and allocation sharing remain unchanged. |
| Unknown/missing names | Resolution and typed `PartNotFound` behavior remain before the hint. |
| Ordered bytes | All output slots are populated by the existing prepared request records in input order. |
| Panic/error ordinal | Private cached helper preserves worker panic conversion and lowest-ordinal selection; caller-facility route is unchanged. |
| Existing unmanaged serial path | Untouched; this candidate is only reached after finite-context parallel admission. |

## Required tests before implementation can be accepted

The following tests are the minimum seam, race, and refusal coverage. They
should be added only with the implementation and then run through the normal
quality gates.

### Cache-state tests

- A cold or missing `EntryId` makes the predicate false.
- A completed retained entry makes it true.
- An entry with an active flight is false, including the provisional-entry
  state during `publish_pending`.
- Duplicate prepared IDs are handled without an allocation and are true when
  the one retained entry is complete.
- Predicate inspection does not change cache diagnostics, the LRU clock,
  `last_used`, retained bytes, eviction count, or budget gauges.

### Public ordered-read tests

- Fresh and primed `read_parts_ordered` return byte-identical ordered batches;
  the primed batch shares each allocation with the retained payload and makes
  no source reads.
- The same finite context records the same `Workers`, `IoConcurrency`,
  scheduler/output `Memory` and `Objects`, and `CpuTasks` behavior before
  and after the operation. All outstanding reservations are released when the
  batch and package are dropped.
- A cache entry evicted after the hint causes a normal cold read and still
  returns exact bytes; it does not cause a refusal or a stale hit.
- A source revision injected after the hint returns `SourceChanged`, and a
  cancellation injected after the hint returns `Cancelled`; neither returns
  a partial batch and clean cache reservations remain intact.
- A private-worker source/provider panic in a cached request returns
  `SourceBackedBatchWorkerPanic` with the same lowest ordinal as the existing
  worker route. Test a one-wave and a reused-wave request shape.

### Explicit facility and refusal tests

- A caller `ScopedWorkers` receives the same per-wave task count and runs the
  same caught tasks on primed input; the hint must not bypass
  `run_all`.
- A zero or one permit `Workers`/`IoConcurrency` budget, scheduler memory or
  scheduler objects one-under boundary, and output memory/objects one-under
  boundary retain the current typed refusal and zero-source-read behavior.
- A `CpuTasks` one-under budget refuses before the hint and leaves no output;
  a successful primed read still consumes the full request count.
- Existing source-cache tests for hit cancellation, exact identity-scoped
  cleanup, eviction, pinning, and failed-load non-retention remain green.

## Measurement and adoption gate

The implementation should first be measured as a paired candidate against the
0786 large primed Parts cases at widths 1/2/4/8/32, with the same affinity,
finite budgets, order blocks, warmups, samples, source observer, and exact
output oracle. Add fresh Parts and small/mixed floor controls to show that the
hint does not move cold or deliberately serial cases. Record cache hit/cold/
flight deltas, source calls, CPU time, whole-child RSS, output hashes, and all
resource snapshots. Do not pool observer timing with native timing.

Adopt only if the paired cached latency improves materially while all refusal,
ordered-byte, source/cancellation, panic, facility, and reservation tests pass.
Any fresh-path or RSS regression, even if the cached p50 improves, must remain
visible in the report. No claim should be generalized from this low-level
warm control to full Office CRUD, physical cold sources, remote ranges, or
whole-program scaling.

## ADR alignment

This candidate is compatible with ADR 0005 because it preserves explicit
budget admission, clean-payload cache eviction/pinning, source identity fences,
and the measured-evidence requirement. It follows ADR 0031 by keeping
`Workers`, `IoConcurrency`, and `CpuTasks` at admission/charge boundaries and
by leaving the caller facility explicit. It follows ADR 0001/0002/0010/0011
because the change stays inside the low-level `litchi-opc` owner and adds no
facade, archive dependency, runtime, or global executor. It follows ADR 0003
because it only reads immutable source-backed state and returns the same
ordered handles. It does not add the derived-value memo or retention shape
covered by ADR 0032; the existing payload cache remains the sole byte
authority.

The design is not an authorization to edit production code. It is a
candidate seam pending implementation, focused tests, paired measurement,
independent review, and the normal commit/cleanup gates.
