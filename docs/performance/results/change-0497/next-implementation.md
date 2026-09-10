# Next implementation after 0497: bounded source-backed OPC multi-part reads

The next production work should be one explicit, bounded multi-Part read path
for `SourceBackedPackage`. It should use the caller's existing
`ExecutionContext` to read independent ZIP Parts concurrently, preserve the
source-backed cache and freshness contract, and return results in request
order. This is the highest-impact open implementation after the 0497 atomic
destination capability. It addresses the still-unmet parallel-execution
requirement in [`docs/GOAL.md`](../../../GOAL.md), rather than extending the
0496 DOCX phase diagnostics.

The [0497 plan](plan.md) and [production review](production-review.md)
establish the atomic filesystem boundary for the bounded DOCX tail append
route; its capture and final gates remain separate evidence work. The
0496 publication percentages cannot justify a production optimization: those
clocks are wall time, and the diagnostic publication path includes the
preallocated output copy and later output hashing in different scopes. The
next change therefore needs a production source operation with its own
before/after oracle and scaling evidence. The full non-iWork scope remains
open for borrowed lifetimes, native producers, cold intersections, arbitrary
append semantics, durable history/composition, and broader CRUD/security.

## Why this is the next measured gap

The goal requires bounded execution with real scaling, a size-derived
threshold, backpressure, and Amdahl analysis. The current source-backed
package already accepts an explicit policy through
`SourceBackedPackage::from_read_at_with_execution_context`, but
`PartView::data` calls the one-Part `read_part` path and `iter_parts` only
produces metadata views. The `PartCache` provides same-Part single-flight and
budgeted retention; it does not provide a cross-Part worker team. The relevant
hooks are [`crates/litchi-core/src/execution.rs`](../../../../crates/litchi-core/src/execution.rs),
[`crates/litchi-opc/src/source_backed.rs`](../../../../crates/litchi-opc/src/source_backed.rs),
and the existing eager-only adapter in
[`crates/litchi-opc/src/execution.rs`](../../../../crates/litchi-opc/src/execution.rs).

The retained substrate measurements make this a real candidate while also
showing where it must refuse parallelism:

| Evidence | Observation | Meaning for the implementation |
| --- | --- | --- |
| [`execution-scaling.json`](../execution-scaling.json), `opc_open_session_scaling` | The six-entry `few-large` OPC case falls from 5.678 ms at one worker to 1.255 ms at twelve (about 4.5x); the 258-entry `many-small` case rises from 0.573 ms to 0.783 ms | Large independent work can scale, while tiny Parts need a sequential threshold |
| [`control-A.json`](../opc-source-cache-release-abba-0100/results/control-A.json) and its reversed/managed siblings | Disjoint source-backed Parts are about 121.2 ms at one worker and 10.2 ms at twelve under the fixed 10 ms source delay; same-Part remains about 10.1 ms at every width | A fixed worker team can expose independent cold loads; same-Part work must retain single-flight semantics |
| [`GOAL_AUDIT.md`](../../GOAL_AUDIT.md) and [`HOTSPOTS.md`](../../HOTSPOTS.md) | The current DOCX lifecycle is serial/one-worker; delayed and short-read providers are large descriptive arms, while source/cache and scaling claims remain open | Measure a source-backed production route across provider boundaries instead of treating a harness pool as proof |

The scaling and cache files are harness/substrate evidence, not a production
latency claim. The new path must prove its own semantic bytes, source counters,
resource release, and worker bounds.

## Minimal production contract

Add an opt-in package-level batch operation, for example
`SourceBackedPackage::read_parts_ordered(&[PackURI]) -> Result<Vec<PartData>>`.
The exact name may follow the crate's API naming, but the contract should be
fixed before implementation:

1. Resolve every requested Part and its declared size before scheduling. Keep
   duplicate requests in the returned order, while allowing the existing
   `PartCache` single-flight state to prevent duplicate cold loads.
2. Use the package's explicit `ExecutionContext`. With no context, one worker,
   one requested Part, or aggregate declared work below
   `ExecutionLimits::min_parallel_bytes`, use the current serial path. For an
   eligible batch, create one local worker/session whose width is at most
   `workers` and whose queued work is bounded by
   `max_in_flight_tasks` and `max_in_flight_bytes`. Do not install a global
   Rayon pool or expose a runtime through ordinary document APIs.
3. Each task must enter the existing `read_part`/`PartCache` path so source
   version fences, cancellation, declared and decoded-size checks, cache
   reservations, and typed ZIP errors remain authoritative. Returned managed
   `PartData` handles retain their existing reservation rules; a parallel call
   must not create an unreserved `Arc<Vec<u8>>` escape.
4. Apply backpressure between batches, check cancellation before dispatch and
   after each completed Part, and join every worker on success or error. A
   failed or cancelled call returns no partial result vector, drops all task
   outputs, and leaves cache flights, in-flight bytes, object reservations,
   and budget counters at the same post-call state as the serial operation.
5. Preserve stable input order and deterministic error selection. A source
   mutation observed before the final result fence, a corrupt member, a read
   limit, or a cancellation must fail closed. Same-Part concurrent reads must
   continue to share one decoded allocation, as covered by
   `concurrent_cold_reads_share_one_archive_load_and_one_arc` in the existing
   source-backed test module.

The existing `OpenSession` and
`soapberry_zip::office::{ParallelReadSession,IndexedArchive::read_many_with_session}`
are implementation hooks for the local bounded scheduler, but the eager
`OpcPackage` route cannot simply be relabeled as source-backed coverage. The
source-backed adapter must integrate cache admission, source fences, and
managed `PartData` ownership. If the generic ZIP session cannot preserve those
invariants, keep the worker adapter in `source_backed.rs` and reuse only its
bounded policy mapping.

The first end-to-end production consumer should be this source-backed batch
operation itself, exercised over a selected multi-Part semantic/content
oracle. Do not change ordinary DOCX APIs or the 0497 atomic publication path
until the low-level operation has passed the cache, cancellation, and error
contract. A later format facade can opt in when it has a genuine independent
set such as media, stories, sections, or validation domains. The serial loop
in `materialize_opc_package_with_accounting` is a useful comparison hook, but
its managed-materialization refusal is an existing ownership boundary and
must not be weakened merely to manufacture a benchmark.

## Required tests and owners

The primary owner is
[`crates/litchi-opc/src/source_backed.rs`](../../../../crates/litchi-opc/src/source_backed.rs),
with focused integration coverage in
[`crates/litchi-opc/tests/source_backed_topology.rs`](../../../../crates/litchi-opc/tests/source_backed_topology.rs)
or a new source-backed batch test module. Add tests for:

- serial equivalence at one worker and below-threshold work, including exact
  bytes, input order, repeated names, and no worker-team creation;
- a gated `ReadAt` source proving at least two independent Parts overlap while
  the observed active task and declared-byte maxima never exceed the chosen
  policy;
- same-Part duplicates sharing one cold load/allocation and disjoint Parts
  producing one deterministic result per request;
- cache capacities below, equal to, and above the working set, with retained
  entries/bytes, in-flight loads, `Resource::Memory`/`Objects`/`Work` usage,
  and all reservations returning to baseline after handles and package drop;
- cancellation during a worker, source-version change during a worker,
  missing/corrupt Part, decoded-size mismatch, and limit failure. Every case
  must join workers, clear flights, drop output handles, and preserve the
  serial typed-error boundary;
- `OwnedSource`, `FileSource`, instrumented/short-read, and delayed `ReadAt`
  providers. A delayed provider must prove overlap with gates and counters,
  not rely on an assumed physical filesystem delay;
- no global runtime creation and no worker count above the caller's local
  policy. Existing same-Part and managed-cache tests remain required gates.

The benchmark owner is
[`tools/perf-baseline/src/lib.rs`](../../../../tools/perf-baseline/src/lib.rs),
reusing the explicit-worker and source-cache evidence helpers near
`run_opc_open_session_scaling` and `run_opc_source_cache_contention`. A
harness-only `OpcCacheWorkerTeam` result is a calibration oracle; it does not
substitute for invoking the production `SourceBackedPackage` operation.

## Bounded measurement

Use the current serial selected-Part loop as the before control and the new
batch API with `workers = 1, 2, 4, 8` (capped by available CPUs) as after
arms. Start with two fixed shapes: the existing `few-large` independent work
set and `many-small` work that should validate the sequential fallback. Add a
realistic multi-Part DOCX/OPC fixture only after the synthetic threshold
invariants pass. Run owned, warm `FileSource`, instrumented/short-read, and
fixed-delay providers; retain source call/byte counters, cache hits/cold
loads/waiters, maximum active tasks, queue batches, lock waits, retained
bytes, budget failures, allocation, RSS, and p50/p95/p99 latency.

The timed read scope starts after source/package setup and fixture preparation
and ends after all requested `PartData` handles arrive. Exact-byte, semantic,
source-version, cache-release, and report serialization oracles run after the
timer. Record scheduler construction separately if the API creates a team per
call; a reused-session service measurement may be a secondary diagnostic, not
the end-to-end claim. Use the repository's normal three warmups, thirty
samples, reversed repeats, and normal/allocator roles once the route is
stable.

Keep two scopes distinct:

- **Read scope:** serial versus bounded source-backed multi-Part reads, with
  source and cache counters attributed to the call.
- **Publication scope:** the existing 0497 sequential/atomic destination
  route, including its required temp-file and sync policy. Do not subtract or
  hide atomic publication cost from a full DOCX operation, and do not claim
  that a read speedup improves publication until a separate full-lifecycle
  capture proves it.

Report fixed-task-count speedup, efficiency, and an Amdahl serial-fraction
estimate only for the `few-large` and realistic independent workloads. Report
the `many-small` rows as threshold/fallback evidence when scheduling loses.
Do not claim cold performance from the existing warm cache rows; a cold
intersection requires its own source-residency and setup protocol.

This change would close one production parallel path and its evidence loop. It
does not complete the broader CRUD checklist, native producer round trips,
borrowed source lifetimes, cold intersections, arbitrary logical-tail append,
durable history/composition, or security coverage.
