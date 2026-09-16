# ADR 0031: I/O concurrency, CPU task and executor budgets in the execution context

- Status: Accepted (2026-09-16, by the owner's decision recorded in [change 0652](../performance/0652-owner-decisions-for-the-third-wave.md); proposed 2026-09-15)
- Date: 2026-09-15
- Supersedes: nothing. Amends: the execution paragraph of
  [ADR 0005](0005-io-memory-and-performance.md) only. Every other paragraph of
  that record, and every other accepted record, is untouched.
- Raised by: [change 0615](../performance/0615-execution-context-completeness-design.md),
  which carries the capability matrix against `docs/GOAL.md` Workstream F, the
  retained scaling evidence, and the coexistence probe whose counts are quoted
  below. Item CORE-3 of
  [change 0587](../performance/0587-remaining-opportunity-survey.md) ranks this
  as a prerequisite rather than a speedup.

## Context

`docs/GOAL.md` rule 9 requires that "CPU parallelism must be opt-in and
controlled by an explicit execution context with thread, memory, I/O,
cancellation, and task-granularity budgets", and Workstream F lists the
capabilities such a context must control: maximum worker count, CPU task
budget, memory and decompression-in-flight budget, I/O concurrency,
cancellation, task size thresholds, and an optional caller-provided executor or
scoped worker facility.

ADR 0005 names a narrower list: "CPU parallelism is opt-in through an execution
context controlling scheduling, affinity, cancellation, thread and memory
budgets. There is no hidden global Rayon pool."

`litchi_core::ExecutionLimits` (`crates/litchi-core/src/execution.rs:38`)
implements the ADR 0005 list and part of the Workstream F list:

| Workstream F capability | Present today | Where |
| --- | --- | --- |
| maximum worker count | yes | `ExecutionLimits::workers` |
| memory and decompression-in-flight budget | yes | `max_in_flight_bytes`, `Resource::Memory` |
| cancellation | yes | `CancellationToken` |
| task size thresholds | partly | `min_parallel_bytes` is an **aggregate** batch threshold; there is no per-task floor |
| CPU task budget | no | `Resource::Work` is charged in declared **bytes**; no dimension counts tasks |
| I/O concurrency | no | nothing bounds concurrent `ReadAt::read_at` calls |
| caller-provided executor or scoped worker facility | no | each session builds its own workers |

The workspace has exactly three explicitly scheduled parallel sessions, and at
`1e4198321` every other `thread::scope` or `rayon` use under `crates/` is test
code:

1. `soapberry_zip::office::ParallelReadSession` — a private Rayon pool built
   **eagerly** in `ParallelReadSession::new` (`office.rs:270-283`), reached only
   through `litchi_opc::OpenSession`;
2. `litchi_cfb::SharedOleBulkRead` — a private Rayon pool built **lazily** on
   the first eligible batch (`shared_bulk.rs:210-225`);
3. `litchi_opc::SourceBackedPackage::read_parts_ordered` — `std::thread::scope`
   workers spawned per operation, reused across waves since change 0499
   (`source_backed/batch.rs:522`, `:585`).

Each honours its own `ExecutionLimits`. Nothing makes them honour each other.
`ExecutionLimits` is `Copy` and is re-applied by every session; `Budget` is the
only part of an `ExecutionContext` that is genuinely **shared**, and it carries
no dimension for workers, for outstanding reads, or for task count. Change
0615's probe drives all three sessions from one hierarchical `Budget` root whose
policy says `workers = W`, and counts what the process actually holds:

| `workers` in the policy | session-owned worker threads at peak | concurrent `read_at` calls, process-wide |
| ---: | ---: | ---: |
| 1 | 0 | 2 |
| 2 | 6 | 4 |
| 4 | 12 | 7–8 |
| 8 | 20 | 7–8 |

The worker figure is `2W + min(W, tasks)` — two persistent Rayon pools plus one
transient scoped wave — and it is identical in all five repeats at every width.
The read figure is the overlap two sessions reading two distinct sources
achieved in a run, against an arithmetic ceiling of 8 (the sum of the two
per-session maxima, each `min(W, independent work)`); it is a lower bound.
Neither number is bounded by anything a caller can set.

This is not a latent risk. It is what a caller who opens an OOXML package with
`OpenSession`, reads a batch of parts, and reads a set of CFB streams in the
same process gets today, from one budget root, with one policy.

## Decision

### 1. The missing budgets belong to the hierarchical `Budget`, not to `ExecutionLimits`

The reason the three sessions do not compose is structural: `workers` lives on
the wrong side of the sharing boundary. `ExecutionLimits` is a per-session
policy value that each session copies and applies independently. `Budget` is a
shared hierarchy in which a parent bounds the sum of its children.

Therefore the three new budgets are added as `Resource` dimensions:

- **`Resource::Workers`** — concurrency permits for OS threads or executor
  slots held by an operation. Peak-denominated: reserved before workers start,
  released when they are joined.
- **`Resource::IoConcurrency`** — permits for outstanding positional reads.
  Peak-denominated, reserved at wave admission.
- **`Resource::CpuTasks`** — cumulative count of scheduled CPU work units
  (decompress, parse, validate). Count-denominated, consumed like
  `Resource::Work` but not in bytes.

`ExecutionLimits::workers` is **retained unchanged** as a per-session ceiling.
A session's effective width becomes `min(limits.workers(), granted Workers
permits)`. Three sessions that share a root granting eight `Workers` permits can
never hold more than eight worker threads between them, whatever their
individual policies say.

### 2. Reservation dimensions fail fast at admission, never mid-operation

`Budget::reserve` refuses; it does not block. A worker that discovered mid-read
that it had no permit would have to abandon work already started, which
`docs/GOAL.md` rule 3 forbids ("never trade a typed refusal for a partial or
guessed edit").

The consumption rule is therefore **admission-time**, matching how
`SchedulerAdmission` already reserves worker stacks and channel state before a
worker starts (`batch.rs:362-470`):

- An operation reserves its `Workers` and `IoConcurrency` permits **before** the
  first task is spawned and before the first payload read.
- If it is granted fewer permits than it asked for, it **narrows the wave** and
  proceeds; it does not fail.
- Every operation that reads at all reserves at least one `IoConcurrency`
  permit, including the serial path. If even one is unavailable, the operation
  refuses with a typed `ResourceLimit { resource: IoConcurrency, observed,
  limit, scope }` **before any source byte is read**, which is the existing
  `Resource::Memory` shape and the existing `ReadResource` refusal identity.
- Permits are released on `Reservation` drop. A session that holds a pool holds
  its `Workers` permits for the pool's lifetime, which makes lazy pool
  construction (§6) load-bearing rather than cosmetic.

A measured permit costs one `fetch_update` CAS per hierarchy level plus an `Arc`
clone into a `SmallVec<[_; 4]>`: 44 ns at depth 1 and 99 ns at depth 3 on this
host (change 0587). It is paid once per wave, not once per read.

### 3. `Limits` gains the three dimensions additively

`Limits::new` is called from 589 sites across 234 files under `crates/` and 22
sites under `tools/`, and `RESOURCE_COUNT` is `6`
(`crates/litchi-core/src/budget.rs:13`). A ninth positional argument would
churn every one of them for no semantic gain, because every one of them is
unbounded in these dimensions **today**.

The proposal is additive:

- `Limits::new` keeps its six arguments and initialises `Workers`,
  `IoConcurrency` and `CpuTasks` to `u64::MAX`. Every existing site therefore
  compiles unchanged and keeps exactly its current meaning.
- A new `const fn Limits::with_execution(self, workers, io_concurrency,
  cpu_tasks) -> Self` sets them.
- `Profile::Server`, `Profile::Desktop` and `Profile::TrustedBatch` **must**
  name finite values, because ADR 0005 requires that "production-safe desktop,
  server, and trusted-batch profiles are finite". Choosing those three numbers
  is a decision this record deliberately leaves to the reviewer; change 0615
  states the evidence that bears on it and declines to invent them.

`Resource` is already `#[non_exhaustive]`, so downstream matches are unaffected;
the 14 in-workspace `match` sites that name `Resource::Depth` as an arm — ten
across streaming and detection paths and their tests, four in `xml-minifier`'s
audit — gain three arms or a wildcard.

### 4. An optional caller-provided scoped-worker facility

`litchi-core` defines the trait and never implements one over a runtime:

```rust
/// A caller-provided facility that runs a bounded set of borrowed tasks.
pub trait ScopedWorkers: Send + Sync + fmt::Debug {
    /// Runs every task exactly once and returns only after all have returned.
    ///
    /// Tasks may run on any thread, in any order, concurrently or serially.
    /// A task never unwinds: callers hand this facility tasks that have
    /// already caught their own panics.
    fn run_all(&self, tasks: &mut [&mut (dyn FnMut() + Send)]);
}
```

Four properties are deliberate:

- **No `'static` bound.** All three sessions hand their workers borrowed state
  (`&SharedOleFile`, `&SourceBackedPackage`, a borrowed archive). A facility
  that required `'static` would force a copy of exactly the bytes the program
  exists to avoid copying. The slice-of-closures shape is the minimum that is
  object-safe, borrow-friendly and implementable over Rayon's `scope`, over
  `std::thread::scope`, and by a serial implementation that just calls each
  task.
- **Run-to-completion.** `run_all` returning means every task has returned. This
  is what both existing implementations already guarantee
  (`pool.install(|| par_iter().collect())`, `thread::scope`), and it is what
  makes cancellation and reservation release deterministic.
- **No unwind contract.** The OPC batch already wraps worker bodies and maps a
  panic to `OpcError::SourceBackedBatchWorkerPanic { ordinal }`
  (`batch.rs:14`, `:576`). Keeping panics on the caller's side of the trait
  means a caller's facility cannot change any error identity.
- **`Debug` supertrait.** `ExecutionContext` is `Clone + Debug`; an
  `Arc<dyn ScopedWorkers>` field must be too.

Attachment is on the context, not on any document API:

```rust
impl ExecutionContext {
    #[must_use] pub fn with_scoped_workers(self, workers: Arc<dyn ScopedWorkers>) -> Self;
    #[must_use] pub fn scoped_workers(&self) -> Option<&Arc<dyn ScopedWorkers>>;
}
```

### 5. Ownership and dependency direction

- The trait lives in `litchi-core`, the most foundational crate
  ([ADR 0002](0002-crate-topology.md)). `litchi-core` gains no dependency: it
  defines a trait and implements nothing.
- `soapberry-zip` does **not** depend on `litchi-core` and must not start to;
  it is a standalone ZIP crate. It defines its own structurally identical
  `ScopedWorkers`, exactly as it already defines `CancellationProbe` for the
  same reason (`office.rs:240`), and `litchi_opc::OpenSession` bridges the two,
  exactly as it already bridges `ExecutionLimits` to `ParallelReadLimits` and
  `CancellationToken` to `CancellationProbe` (`litchi-opc/src/execution.rs:34-49`).
  No new dependency edge is created anywhere.
- Rayon remains a private implementation detail of `litchi-cfb` and
  `soapberry-zip`. No Rayon or Tokio type appears in any public signature. There
  is still no global pool.
- `docs/GOAL.md` rule 7 forbids exposing "executors, runtime handles" through
  **ordinary public CRUD APIs**. `ScopedWorkers` is a caller-implemented trait
  that appears only on `ExecutionContext`, which is already an advanced-ingress
  type that no `Document`, `Workbook`, `Presentation`, worksheet, slide or
  paragraph signature names. Acceptance of this record is conditional on an
  audit proving that remains true (§Verification).

### 6. How the three existing sessions consume it

| Session | Change |
| --- | --- |
| `soapberry_zip::office::ParallelReadSession` | Build the pool **lazily**, on the first batch that qualifies for parallel execution, as `SharedOleBulkRead` already does. This alone removes the eager threads change 0615 measured at `threads_after_zip_session_new` from a session that may never take its parallel branch. Reserve `Workers` permits when the pool is built and hold them for its lifetime; take the caller's facility instead of a pool when one is attached. |
| `litchi_cfb::SharedOleBulkRead` | Reserve `Workers` when the lazy pool is built; reserve `IoConcurrency` for the batch width at `read_streams` batch admission, before `read_one`; take the caller's facility when attached. `Resource::CpuTasks` is consumed once per stream request. |
| `litchi_opc::SourceBackedPackage::read_parts_ordered` | Add `Workers` and `IoConcurrency` to the existing `SchedulerAdmission::reserve`, which already reserves worker stacks, channel endpoints and objects before startup. Narrow the wave to the granted width instead of failing. Consume `CpuTasks` once per prepared request. Replace the scoped spawn with the caller's facility when one is attached; keep `thread::scope` otherwise. |

Every one of these is an addition inside an existing admission step. None moves
a fence, changes an error identity, changes output bytes, or changes what any
existing caller observes when the new dimensions are `u64::MAX` — which they
are for every existing caller (§3).

### 7. A per-task size floor

`min_parallel_bytes` is an **aggregate** threshold: a batch of sixty-four 16 KiB
parts totals 1 MiB and clears any realistic value, yet change 0498 measured that
batch at 94.8 µs serial against 560.0 µs at four workers, and change 0009
measured 258-task OPC and 256-stream CFB many-small batches at 0.73× and 0.52×.
The policy set therefore cannot today express the one threshold the retained
evidence most clearly justifies.

`ExecutionLimits` gains `min_task_bytes: u64`, a per-task floor validated
against `max_in_flight_bytes` in the same way `min_parallel_bytes` already is.
A batch whose individual tasks are below the floor stays serial even when the
aggregate clears `min_parallel_bytes`.

This field is not named by survey item CORE-3 or by Workstream F's list; it is
proposed because the measured evidence in change 0615 points at it more
directly than at anything else in this record. A reviewer who wants the record
narrowed should strike §7 first.

### 8. What remains opt-in

Everything.

- There is no default `ExecutionContext` and no default `Budget`. Constructing
  one remains entirely the caller's act.
- Absence of an attached facility means the existing private-pool behaviour.
- The three new dimensions default to `u64::MAX` through `Limits::new`, so an
  existing caller observes no change at all.
- No ordinary CRUD method starts parallel work. The three sessions remain
  low-level, opt-in APIs with no format-crate or facade caller
  (`read_parts_ordered`, `bulk_read` and `OpenSession` have zero callers outside
  `litchi-opc`, `litchi-cfb` and the harness at `1e4198321`).
- There is still no hidden global Rayon pool, and this record creates no path to
  one.

## Consequences

**Gained.** A caller can, for the first time, bound the worker threads, the
outstanding positional reads and the scheduled task count of a *process* rather
than of one session, by giving the three sessions children of one budget root.
A caller can supply its own thread pool instead of accepting three private ones.
A caller on a latency-bearing or metered source can cap concurrent reads without
capping CPU width.

**Paid.** Three `Resource` variants, `RESOURCE_COUNT` 6 → 9, three arms at each
of the 14 exhaustive `Resource` matches, one new `Limits` constructor, one new
public trait in `litchi-core`, one mirrored trait in `soapberry-zip`, one bridge
in `litchi-opc`, and an admission change in each of the three sessions. Three
finite numbers must be chosen for each of the three production profiles. Every
`Budget` node grows by three `AtomicU64`.

**Not claimed.** This record makes no performance claim. It does not make any
path faster; it makes an existing path bounded. Change 0615 carries
`performance_claim: none` and registers nothing.

**Rejected alternatives.**

- *Put `io_concurrency` and `cpu_tasks` on `ExecutionLimits` beside `workers`.*
  Cheaper, and it fixes nothing: the measured failure is that per-session policy
  values do not compose, and adding more per-session policy values would produce
  the same table at a higher field count.
- *A process-global registry of live sessions.* Rejected on ADR 0005 grounds
  ("no hidden global Rayon pool" is a specific case of a general refusal of
  ambient process state) and on `docs/GOAL.md` rule 8.
- *A blocking semaphore instead of a fail-fast reservation.* Rejected: a
  blocking acquire inside a worker that already holds a `Workers` permit is a
  deadlock the type system cannot rule out, and a mid-operation stall is not a
  budget behaviour any existing `Resource` has.
- *Make the facility `FnOnce`-based or `'static`.* Rejected: every existing
  session hands its workers borrowed package and source state.

## Verification

Acceptance should require, at minimum:

1. A test that three sessions sharing one budget root with `Workers: N` hold at
   most `N` worker threads between them, counted from `/proc/self/task` — the
   change 0615 probe, converted into a test and re-run against the
   implementation.
2. A test that an `IoConcurrency` grant of one forces every session onto its
   serial path and that the resulting bytes are identical.
3. A test that a zero-permit `IoConcurrency` refusal happens **before** the
   first source read, with zero `read_at` calls observed, mirroring the existing
   one-byte-under `Memory` boundary case in change 0088.
4. A test that a caller-supplied `ScopedWorkers` that runs every task serially
   produces byte-identical results to the private-pool path, and that a caller
   facility never changes an error identity, an ordinal selection, or a
   cancellation outcome.
5. An API audit proving no ordinary CRUD signature names `ScopedWorkers`,
   satisfying `docs/GOAL.md` rule 7.
6. A crate-boundary check proving `soapberry-zip` still does not depend on
   `litchi-core`.
7. Re-running the change 0009 and 0498/0499 scaling captures with the new
   dimensions set to `u64::MAX`, showing no movement outside the A/A floor — the
   proof that the opt-in default costs nothing.
8. A decision, with stated reasoning, on the three finite profile values.

Nothing in this record may be implemented before a human accepts it, and an
implementation must be measured before any of it is claimed.
