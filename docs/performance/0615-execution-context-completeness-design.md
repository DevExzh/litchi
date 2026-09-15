# 0615: the execution context bounds one session at a time, and the workspace runs three

Status: design only, plus a proposed ADR for human review. No production change
and `performance_claim: none` — this record carries deterministic thread and
positional-read counts, a capability matrix and a set of admission gates. **No
timing is measured and none is claimed.**

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

## Why this record exists

`docs/GOAL.md` rule 9 requires that "CPU parallelism must be opt-in and
controlled by an explicit execution context with thread, memory, I/O,
cancellation, and task-granularity budgets", and its Workstream F names seven
capabilities such a context must control. Item CORE-3 of
[change 0587](0587-remaining-opportunity-survey.md) ranks the gap 35th and calls
it "a prerequisite rather than a speedup":

> `ExecutionLimits` […] lacks an I/O concurrency limit (nothing bounds
> concurrent `read_at`s issued from workers), a CPU task budget distinct from
> byte-denominated `Work`, and a caller-provided executor or scoped-worker
> facility — each of the three sessions builds its own pool […] so three pools
> can coexist in one process with no shared cap.

That survey entry names the gap from the source. This record measures it, finds
that the mechanism is worse than "three pools can coexist" in one specific way
that changes the design, and freezes both the design and the gates any
implementation must clear. The design itself is
[proposed ADR 0031](../adr/0031-execution-context-budgets.md), which is **not
accepted** and is deliberately absent from the accepted table.

Three results, in order of how much they change the picture:

1. **The problem is not pool duplication; it is that `workers` is on the wrong
   side of the sharing boundary.** `ExecutionLimits` is `Copy` and every session
   re-applies it independently. `Budget` is the only shared part of an
   `ExecutionContext`, and it has no dimension for workers, outstanding reads or
   task count. So the fix is not "share a pool"; it is "make the missing budgets
   hierarchical resources". Measured: one budget root, one policy saying
   `workers = W`, and **2W + min(W, tasks)** worker threads in the process.
2. **Every one of the three sessions is unreachable from any format crate or
   from the facade.** At `1e4198321`, `read_parts_ordered`, `bulk_read` and
   `OpenSession` have zero callers outside their own crate's tests, one
   `litchi-opc` example, and the harness. Every `thread::scope`, `rayon` and
   `thread::spawn` under `crates/` other than the three sessions is test code.
   So the coexistence problem is real for a caller who opts in three times — and
   the *independent work* Workstream F enumerates has no parallel path at all.
3. **The policy set cannot express the one threshold the retained evidence most
   clearly justifies.** `min_parallel_bytes` is an aggregate: sixty-four 16 KiB
   parts total 1 MiB and clear any realistic value, and change 0498 measured
   that batch at 94.8 µs serial against 560.0 µs at four workers. A per-task
   floor is missing, and no record proposes one.

## 1. The capability matrix against Workstream F

`docs/GOAL.md` Workstream F asks for an execution context that can control
seven things. `litchi_core::ExecutionLimits`
(`crates/litchi-core/src/execution.rs:38-44`) plus `ExecutionContext`
(`:196-260`) and `Budget` (`crates/litchi-core/src/budget.rs:18-25`) provide
four and a half of them.

| Workstream F capability | State | Where, at `1e4198321` |
| --- | --- | --- |
| maximum worker count | **present, per session** | `ExecutionLimits::workers` (`execution.rs:39`). Validated against `max_in_flight_tasks`; `AffinityPolicy::Inherit` is the only policy. |
| memory and decompression-in-flight budget | **present** | `max_in_flight_bytes` (`:41`) plus `Resource::Memory` reserved per batch (`shared_bulk.rs:151`, `batch.rs:307-330`). |
| cancellation | **present** | `CancellationToken` (`execution.rs:171-186`); checked before scheduling, between batches, before and after each member read. |
| task size thresholds | **half present** | `min_parallel_bytes` (`:42`) is an **aggregate** batch threshold. There is no per-task floor. See §5. |
| CPU task budget | **absent** | `Resource::Work` is charged in declared **bytes** (`shared_bulk.rs:138` charges the whole request's stream bytes). No `Resource` counts tasks. |
| I/O concurrency | **absent** | Nothing bounds concurrent `ReadAt::read_at`. Measured in §3. |
| optional caller-provided executor or scoped worker facility | **absent** | Each session constructs its own workers. `litchi-core` defines no such trait. |

Two further Workstream F requirements have no implementation at all:

- **"Use bounded pipelines where I/O, decompression, parsing, semantic
  processing, and serialization can overlap."** Every session is
  wave-synchronous. `soapberry-zip` flushes a batch and collects it before
  starting the next (`office.rs:396-421`); the OPC batch's contract from change
  0499 is explicit that "later waves cannot begin until the current wave has
  completed successfully". Backpressure is by wave, not by a bounded queue, so
  no two phases ever overlap.
- **"Apply backpressure so parallel decompression cannot exceed the memory
  budget."** This one *is* satisfied, by `max_in_flight_bytes` and the per-batch
  `Resource::Memory` reservation — but per session, with the same composition
  hole as `workers`.

## 2. What independent work has no parallel path today

Workstream F enumerates the dependency DAG worth building. Measured against the
tree at `1e4198321`:

| Independent work Workstream F names | Parallel path today | Reachable from a format crate or the facade |
| --- | --- | --- |
| separate ZIP entries | `ParallelReadSession` (eager OPC open) and `read_parts_ordered` (source-backed batch) | **no** — `OpenSession` and `read_parts_ordered` have zero callers outside `litchi-opc`'s own tests, one example, and the harness |
| separate CFB streams | `SharedOleBulkRead` | **no** — `bulk_read` has zero callers outside `litchi-cfb`'s own tests and the harness |
| independent worksheets, slides, stories, sections | **none** | n/a — every `rayon`/`thread::scope`/`thread::spawn` in `litchi-xlsx`, `litchi-xlsb`, `litchi-docx`, `litchi-pptx`, `litchi-xls`, `litchi-doc`, `litchi-ppt`, `litchi-ooxml-common`, `litchi-drawingml`, `litchi-ole-common` and `litchi` is test code |
| disjoint edit subplans | **none** | `JoinedSubEdits::join` is a serial O(n²·k) scan (`litchi-core/src/patch.rs:1755-1790`) |
| independent validation domains | **none** | no site |
| compression of changed output members | **none** | zero `rayon`/`scope` hits under `litchi-opc/src/pkgwriter.rs`, `atomic.rs`, the `litchi-cfb` writers and the soapberry-zip writers — this is item CORE-4 / SAVE-6, which needs its own frozen design |

The consequence for the ranking is worth stating plainly: **completing the
execution context is a prerequisite for work that does not yet exist.** No
measured end-to-end scenario today admits more than one session, because no
end-to-end scenario admits even one. The 0587 survey's own falsification
criterion for CORE-3 — "falsified if no measured scenario ever admits more than
one session at a time (then pool duplication is theoretical)" — is therefore
*satisfied for today's callers* and the item survives only as a prerequisite.
This record does not claim otherwise, and §7 makes the ordering explicit.

## 3. Measured: the coexistence counts

A scratch probe (`results/change-0615/probe/`) drives all three sessions from
**one** hierarchical `Budget` root, with one `ExecutionLimits` value of
`workers = W`, over an in-memory four-member OPC archive and a four-stream CFB
file, one MiB per member. It counts live OS threads from `/proc/self/task` at
each stage and from inside a worker's own `read_at`, and counts simultaneous
`read_at` calls with an in-flight atomic.

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, `rustc 1.95.0`,
`cargo build --release`, pinned with `taskset -c 10`, eight agents active on the
host. These are **counts**, not timings; the host's load does not move them.
Five repeats were taken at each width (`results/change-0615/probe-repeats.txt`);
every thread count below is **identical in all five**.

| `workers` | threads after `OpenSession::new` | after CFB bulk read | peak during OPC part batch | peak, all three live | session-owned worker threads at peak |
| ---: | ---: | ---: | ---: | ---: | ---: |
| 1 | 1 | 1 | 1 | 4 | **0** |
| 2 | 3 | 5 | 7 | 10 | **6** |
| 4 | 5 | 9 | 13 | 16 | **12** |
| 8 | 9 | 17 | 21 | 24 | **20** |

The baseline is one thread. The "all three live" column is sampled from inside a
worker's `read_at` while a CFB bulk read, an OPC part batch and an eager OPC
open run concurrently, and includes the probe's own main thread and three driver
threads; the last column subtracts those four.

The figure is exactly `2W + min(W, tasks)`: two **persistent** Rayon pools, one
per pool-owning session, plus one **transient** scoped wave whose width the OPC
batch caps at `min(workers, max_in_flight_tasks, requests.len())`
(`batch.rs:174-180`), which is why `W = 8` yields 20 rather than 24. At `W = 1`
all three sessions decline to go parallel — `ParallelReadSession::new` builds no
pool, `SharedOleBulkRead` requires `workers > 1`, and the OPC batch requires
`workers > 1` — so the count is zero. That is the one configuration in which the
current design composes, and it composes by not being parallel.

Concurrent positional reads, with a fixed 500 µs per-read delay as a
deterministic instrument that widens the observation window (the same role
change 0088's fixed 10 ms source delay plays; it models no production storage):

| `workers` | CFB bulk read | OPC part batch | process-wide, both sources at once |
| ---: | ---: | ---: | ---: |
| 1 | 1 | 1 | 2 |
| 2 | 2 | 2 | 4 |
| 4 | 4 | 4 | 7–8 |
| 8 | 4 | 4 | 7–8 |

The per-session columns are `min(W, independent work)` in all five repeats at
every width: four CFB streams and four OPC parts cap both at four. The
process-wide column is a **lower bound** on what is reachable — it is the largest
overlap the two independently scheduled sessions' read windows happened to
achieve in a run — and it reaches the arithmetic sum of the two per-session
maxima, 8, in three of the five repeats at `W ≥ 4`. A sixth, separately retained
single run (`probe-w4-delay500.txt`) recorded 3 rather than 4 for the CFB
session, which is what a lower-bound observation does and is retained rather
than discarded.

Nothing a caller can set bounds either column. `ExecutionLimits` has no field
for it, and `Resource` has no dimension for it.

Without the delay instrument, the same run reports per-session concurrency of 1:
a `read_at` over an in-memory `Vec` is a `memcpy`, and on one pinned core four
workers' copies do not overlap. The thread counts are identical in both legs
(`probe-w4-nodelay.txt` against `probe-w4-delay500.txt`), which is why the
thread table above is the load-bearing evidence and the read table is the
supporting one.

Two negative findings, both retained because they close questions:

- **The pools do not leak.** `threads_after_sessions_dropped` still reports
  `2W + 1` at the instant every session is dropped, but a sample 250 ms later
  reports 1 at every width. Rayon terminates a dropped pool's workers
  asynchronously; the threads are reaped. The problem is lifetime, not leakage:
  a session that is *held* holds its threads, and `ParallelReadSession` holds
  them from construction whether or not any batch ever qualifies.
- **The eager pool is measurable.** `threads_after_zip_session_new` is already
  `W` before a single byte is read, because `ParallelReadSession::new` builds
  its pool in the constructor (`office.rs:270-283`), unlike `SharedOleBulkRead`,
  which builds lazily on the first eligible batch (`shared_bulk.rs:210-225`).
  Making the ZIP session lazy is the smallest independently useful piece of this
  design, and it needs no ADR.

## 4. The retained scaling evidence

No new scaling measurement was taken for this record. The retained evidence is
what any implementation must beat, and it is consistent across five records and
three years of the program's own numbering.

**Change 0009** (12 visible CPUs, caller-created local pools, p50 ms,
incompressible corpora). Efficiency is speedup divided by width; the serial
fraction is `s = (1/S − 1/N) / (1 − 1/N)` at `N = 12`.

| Case / exact work | w1 | w2 | w4 | w8 | w12 | S12 | E12 | Amdahl `s` |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| OPC few-large, 6 ZIP tasks / 16,778,178 B | 5.678 | 4.015 | 1.304 | 1.260 | 1.255 | 4.52× | 37.7% | ≈ 15.0% |
| CFB few-large, 4 streams / 16,777,216 B | 3.988 | 0.887 | 0.696 | 0.742 | 0.672 | 5.93× | 49.4% | ≈ 9.3% |
| OPC many-small, 258 tasks / 285,282 B | 0.573 | 0.759 | 0.674 | 0.693 | 0.783 | **0.73×** | 6.1% | out of model |
| CFB many-small, 256 streams / 262,144 B | 0.201 | 0.182 | 0.151 | 0.210 | 0.386 | **0.52×** | 4.3% | out of model |

0009's own caveats are load-bearing and are repeated here rather than dropped:
CFB w2/w4 and OPC w4 are **superlinear** in this warm-memory measurement (E of
224.8%, 143.2% and 108.9%), so their negative formula outputs are not meaningful
Amdahl estimates; the many-small `S < 1` outputs violate the model in the other
direction. Large-task scaling saturates at the available independent work — six
OPC members and four CFB streams — so the w8 and w12 columns measure scheduling,
not width. And the tails do not justify a wide default: at 12 workers OPC
few-large is 1.255 / 2.674 / 2.889 ms p50/p95/p99.

**Change 0498** (source-backed OPC part batches, pooled median µs, one frozen
executable):

| Corpus / source | Serial | Batch 1 | Batch 4 | Batch 8 |
| --- | ---: | ---: | ---: | ---: |
| Four 1 MiB parts / owned | 575.9 | 581.7 | 82.4 | 82.5 |
| Four 1 MiB parts / warm file | 578.6 | 577.2 | 91.0 | 91.7 |
| Four 1 MiB parts / short-read delay | 82,119.6 | 82,474.8 | 20,028.2 | 20,016.4 |
| Sixty-four 16 KiB parts / owned | 94.8 | 93.9 | **560.0** | **480.6** |
| Sixty-four 16 KiB parts / warm file | 182.3 | 201.8 | **577.2** | **512.8** |
| Sixty-four 16 KiB parts / short-read delay | 29,572.8 | 29,600.1 | 8,016.4 | 4,190.6 |

0498 retains 50 adverse flags on the aggregate comparison and says the
many-small regressions "are material" and that "per-wave thread creation is a
plausible contributor, but the measurements do not isolate its cost". It also
warns that the few-large owned medians "appear superlinear at four workers" and
"must not be fitted to a negative or clamped Amdahl serial fraction".

**Change 0499** isolated most of that thread-creation cost: operation-local
worker reuse took many-small owned batch-4 from 542.70 to 186.04 µs (−65.72%)
and cut traced `clone3` calls from 64 to 4 at width four. The conclusion that
matters for the gates below is the one 0499 states itself: even after the fix,
**186.04 µs against an after-serial p50 of 95.86 µs** — the parallel route is
still 1.94× slower than serial for those parts. It also retains an unresolved
tail regression: the delayed-provider batch-8 aggregate p99 rises 30.79% despite
a 3.67% median reduction.

**Change 0088** ran a 62-record control-versus-managed contention matrix on a
256-part corpus at widths 1/2/4/8/12 and accepted **no** speedup: "zero cells
passed both directional confidence gates". **Change 0240** fixed what the
harness may say about parallelism — the `parallel_metrics` envelope is marked
`claim: "descriptive"`, "process thread count and lock wait remain unavailable",
and `waiter_joins` is never converted into time. **Change 0405** added a lock
observation seam and stated its own boundary: "its lock nanoseconds therefore
cannot be compared with an uninstrumented run as a pure contention or `Condvar`
wait measure."

Taken together the five records say one thing: on this corpus class, width helps
only when the independent work is few-and-large, it hurts when the work is
many-and-small, and every number that looked like a scheduler win has been
withdrawn or fenced by its own record.

## 5. What the evidence says that the survey item does not

`min_parallel_bytes` is validated against `max_in_flight_bytes` and compared
against the **aggregate** batch size in all three sessions (`office.rs:396`,
`shared_bulk.rs:160-162`, `batch.rs:182-188`). Sixty-four 16 KiB parts total
1 MiB. Any `min_parallel_bytes` a caller would plausibly choose is cleared by
that batch, and 0498 measured it at 5.91× slower than serial at four workers.

So the policy set has a hole that neither CORE-3 nor Workstream F's list names:
there is no **per-task** size floor. Proposed ADR 0031 §7 adds
`ExecutionLimits::min_task_bytes` for it, and flags it as the part of that
record a reviewer should strike first, because it is the one element not traced
to a stated requirement — only to a measurement.

## 6. Why the design is what it is

The full argument is in [proposed ADR 0031](../adr/0031-execution-context-budgets.md).
The three decisions that the measurement, rather than the survey, forced:

1. **The new budgets are hierarchical `Resource` dimensions, not
   `ExecutionLimits` fields.** Adding `io_concurrency` and `cpu_tasks` beside
   `workers` would be cheaper and would change nothing: §3 shows that the
   failure is that per-session policy values do not compose. `Resource::Workers`,
   `Resource::IoConcurrency` and `Resource::CpuTasks` compose because `Budget`
   already does.
2. **Permits are reserved at admission and never mid-operation.**
   `Budget::reserve` refuses rather than blocks, and a worker that discovered
   mid-read that it had no permit would have to abandon started work, which
   `docs/GOAL.md` rule 3 forbids. An operation granted fewer permits than it
   asked for narrows its wave; an operation that cannot get even one
   `IoConcurrency` permit refuses **before its first source byte**, which is the
   existing `Resource::Memory` refusal shape and the boundary change 0088
   already tests one byte under.
3. **`Limits::new` keeps six arguments.** It has 589 call sites across 234 files
   under `crates/` and 22 under `tools/`, every one of which is unbounded in
   these dimensions today. The dimensions default to `u64::MAX` and a separate
   `Limits::with_execution` sets them, so every existing caller keeps exactly its
   current meaning and the opt-in stays opt-in. The production profiles must
   name finite values, because ADR 0005 requires them to be finite; ADR 0031
   deliberately leaves those three numbers to the reviewer.

## 7. Admission gates for any implementation

No part of this design may be implemented before a human accepts ADR 0031. Once
accepted, an implementation must clear all of the following, in this order. The
order matters: gates 1 and 2 are free and independently useful, gate 3 is the
prerequisite, and gates 4 to 7 are what any *use* of the parallelism must show.

**G1 — the lazy ZIP pool, separately.** Make `ParallelReadSession` build its
pool on the first qualifying batch, as `SharedOleBulkRead` already does. Gate:
the probe's `threads_after_zip_session_new` falls from `W` to 0 at every width,
and every `opc_open_session_scaling` harness row is unmoved outside the A/A
floor. No ADR is needed for this and it should not wait for one.

**G2 — the opt-in default costs nothing.** With all three new dimensions at
`u64::MAX`, re-run the change 0009 `opc_open_session_scaling` and
`cfb_bulk_read_scaling` captures and the change 0498/0499 batch captures.
Gate: no p50, p95 or p99 moves outside its A/A floor in the same window, and the
retained per-cell counters — cache admissions, cold loads, flights, waiters,
source calls, `Work` and `InputBytes` — are identical.

**G3 — the budgets actually compose.** Convert `results/change-0615/probe/` into
a test. Gates, at widths 1, 2, 4 and 8:
- three sessions sharing one root with `Resource::Workers = N` hold at most `N`
  session-owned worker threads between them, counted from `/proc/self/task`;
- an `IoConcurrency` grant of 1 forces every session onto its serial path and
  produces byte-identical output;
- a zero-permit `IoConcurrency` refusal is typed, names the resource, observed,
  limit and scope, and happens with **zero** `read_at` calls observed;
- a caller-supplied `ScopedWorkers` that runs every task serially produces
  byte-identical output and changes no error identity, no ordinal selection and
  no cancellation outcome.

**G4 — scaling, stated in full.** For each scenario, report p50/p95/p99 at
widths 1, 2, 4, 8 and N, with speedup, efficiency and the Amdahl serial fraction
at each width, and classify every cell as baseline, valid, superlinear, slowdown
or out-of-model in the shape change 0088's harness already emits. A superlinear
or `S < 1` cell is reported as out-of-model, never fitted. Efficiency below 25%
at the widest width is a review trigger, not a result to average away.

**G5 — no regression for small tasks.** The many-small cases are the ones five
records agree get worse. Gate: with the design's thresholds set as the
implementation recommends, the 64 × 16 KiB and 256-stream corpora are **not
slower than serial at any width**, at p50 and at p99. Change 0499's residual —
186.04 µs parallel against 95.86 µs serial — is the number to beat, and if a
per-task floor is what achieves it by keeping those batches serial, that is a
pass, because the gate is the end-to-end result and not the width.

**G6 — contention counted, not timed.** Report acquisition counts through the
change 0405 seam, `clone3` counts in the shape change 0499 traced, and context
switches and CPU migrations per operation. Do not convert waiter counts into
time (change 0240) and do not compare instrumented lock nanoseconds with an
uninstrumented run (change 0405). A rise in context switches or migrations
without a matching p50 improvement is a review trigger.

**G7 — the boundary audits.** No ordinary CRUD signature names `ScopedWorkers`
(`docs/GOAL.md` rule 7); `soapberry-zip` still does not depend on `litchi-core`
(ADR 0002); no global pool exists (ADR 0005); `#![forbid(unsafe_code)]` holds
where it holds today.

**Falsification.** The design is falsified if G3 cannot be met without a
blocking acquire — that is, if narrowing a wave to the granted width turns out
to change an error identity or an ordinal selection somewhere. It is weakened,
but not falsified, if a reviewer strikes ADR 0031 §7: the composition argument
does not depend on the per-task floor.

## 8. ADR compliance

| Record | Reading |
| --- | --- |
| [ADR 0001](../adr/0001-priorities-and-api-layers.md) | `ScopedWorkers` is a caller-implemented trait on `ExecutionContext`, an advanced-ingress type. It is not an `Arc<RwLock<T>>`, a source generic, a runtime handle or a package ID in a normal document signature. Gate G7 proves it. |
| [ADR 0002](../adr/0002-crate-topology.md) | No new dependency edge. `litchi-core` defines a trait and implements nothing; `soapberry-zip` mirrors it exactly as it already mirrors `CancellationProbe`, and `litchi-opc` bridges, exactly as it already bridges `ExecutionLimits` and `CancellationToken` (`litchi-opc/src/execution.rs:34-49`). |
| [ADR 0003](../adr/0003-snapshots-edits-and-patches.md) | Untouched. Nothing here reaches snapshots, edits, commits, patches or conflicts. `JoinedSubEdits` is named in §2 as work with no parallel path, not as work to parallelize. |
| [ADR 0005](../adr/0005-io-memory-and-performance.md) | **This is the record that would be amended**, and only its execution paragraph: "scheduling, affinity, cancellation, thread and memory budgets" gains I/O concurrency, a CPU task budget and an optional caller facility. Its "production-safe desktop, server, and trusted-batch profiles are finite" clause becomes a requirement on the three new dimensions. "There is no hidden global Rayon pool" is strengthened, not weakened: the design creates no path to one and gate G7 checks it. Because this is an amendment to an accepted record, `docs/GOAL.md` line 80 applies and a separate proposed ADR is drafted for human review rather than implemented. |
| [ADR 0006](../adr/0006-validation-security-and-compatibility.md) | Untouched. No validation moves, no refusal is traded for a partial result, and the new refusal (`IoConcurrency` exhausted) fires before any read, which is stricter than any existing behaviour rather than weaker. |
| [ADR 0010](../adr/0010-facade-archive-ownership.md) / [ADR 0011](../adr/0011-ooxml-physical-package-ownership.md) | Untouched. No ownership moves; every change is inside an existing admission step of the crate that already owns that work. |

No accepted ADR is weakened and none is invoked as permission. One proposed ADR
is drafted, exactly as `docs/GOAL.md` requires when an otherwise-desirable
change would amend an accepted record.

## 9. Validation preserved

Nothing under `crates/` changed, so every validation, refusal, limit and fence
is bit-for-bit what it was at `1e4198321`. The probe is a separate Cargo project
with path dependencies that reads only the public API; it adds no `unsafe`, no
feature flag and no test hook, and it is not part of the workspace.

## 10. Limitations — what is not claimed

- **No timing, no speedup, no regression, and no claim.**
  `performance_claim: none`; nothing is registered in the claim registry.
- The thread and read counts are **counts on one host, from one probe, on
  synthetic in-memory corpora** of four 1 MiB members. They are deterministic
  and were stable across every run, but they measure the *structure* of the
  three sessions, not any real document workload.
- The process-wide concurrent-read figure is a **lower bound**: it is the
  largest overlap two sessions' read windows achieved in one run, not a proof of
  the maximum reachable.
- The 500 µs per-read delay is a **deterministic instrument**, not a model of
  any storage device, network filesystem or provider. Without it the same probe
  reports per-session read concurrency of 1 on one pinned core.
- The probe runs on Linux and reads `/proc/self/task`. Nothing here says what a
  macOS or Windows runtime adapter would report.
- All scaling numbers in §4 are **retained** from changes 0009, 0498 and 0499.
  They were taken on different hosts, different corpora and different tree
  states, and none was re-measured here. Their own out-of-model caveats apply
  and are reproduced rather than dropped.
- The design in ADR 0031 is **proposed, not accepted**. Nothing in it may be
  implemented, and no code may cite it, until a human accepts it. The three
  finite profile values it requires are deliberately not chosen here.
- Gate G1 (the lazy ZIP pool) is the only part that needs no ADR, and it is not
  implemented here either: this record's scope is documentation plus the probe.
- Whether any of this is *worth* doing for a real workload is unknown, and §2
  says why: no format crate or facade path reaches any of the three sessions, so
  there is no end-to-end scenario to measure yet. Item CORE-4 / SAVE-6 (parallel
  deflate of changed members) is the most likely first real consumer and needs
  its own frozen design first.

## Retained evidence

[`results/change-0615/README.md`](results/change-0615/README.md) — the probe
source and its `Cargo.toml`, the five raw probe outputs (widths 1, 2, 4 and 8
with the delay instrument, and width 4 without it), the five-repeat table
`probe-repeats.txt`, `gates.txt`, `decision.json` and `log-sections.md`.
