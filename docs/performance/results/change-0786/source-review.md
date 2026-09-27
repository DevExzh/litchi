# 0786 source review

This is a source and harness review for the bounded 0786 read-session batch. The
production tree is unchanged at `76dde712f53f007d51be6b29a64b575c3a4c3eb9`.
This review did not run Cargo, native captures, profilers, or benchmarks.

The program requirement is still evidence, rather than capability: `docs/GOAL.md`
requires cold/warm/concurrent workloads (`:27-43`), p50/p95/p99, source and
range-read accounting, and scaling at 1/2/4/8/available widths (`:206-230`).
Its parallelism workstream requires explicit workers, CPU tasks, memory, I/O,
cancellation, and task thresholds (`:568-606`), and the definition of done
requires real bounded scaling evidence (`:799-817`). The CRUD checklist says
that capability does not establish a semantic claim (`docs/CRUD_Scenario_Checklist.md:3-12`),
and requires caller-selected limits and reproducible performance evidence
(`:71-81`).

The current receipts explain why this batch is useful but bounded:

* `docs/performance/0676-execution-budget-composition.md:3-47` and
  `docs/performance/results/change-0676/README.md:7-10,31-51` establish the
  three low-level read routes and finite resource-admission behavior, but
  explicitly retain no latency, throughput, or scaling table.
* The current baseline exposes only `opc_open_session_scaling` and
  `cfb_bulk_read_scaling` (`tools/perf-baseline/src/lib.rs:3095-3097`). Its
  `execution_context` uses `Limits::new` (`:25731-25768`), while
  `Limits::new` leaves Workers, IoConcurrency, and CpuTasks unbounded unless
  `with_execution_io` is applied (`crates/litchi-core/src/budget.rs:73-133`).
  The baseline result retains only requested worker and logical task/byte
  fields (`tools/perf-baseline/src/lib.rs:60517-60533`), and its process-wide
  thread count is deliberately unavailable (`tools/perf-baseline/src/parallel_metrics.rs:195-203`).
* Therefore this packet can provide a finite-budget, low-level scaling receipt;
  it cannot promote a CRUD row or close physical-cold, range-source, native
  producer, or end-to-end concurrency requirements. That boundary is also
  stated in `docs/performance/CRUD_COVERAGE.md:3-13,91-110`.

## OPC fixture and `load_parts_eager`

The standalone fixture is explicit in `tools/perf-execution/src/main.rs`:

* `MEMBER_COUNT` is 32 and member sizes are 4 KiB, 256 KiB, or mixed 31 large
  plus one 4 KiB (`:25-34,68-77`). The corpus builder creates 32 payloads and
  hashes each one (`:391-401,436-457`).
* `build_opc` writes exactly two structural ZIP members first,
  `[Content_Types].xml` and `_rels/.rels`, then the 32
  `custom/member00.bin` through `custom/member31.bin` payload members
  (`:403-423`). Thus “all32 payload” means 32 ordinary payload members; it
  does not mean 32 total ZIP entries. The OPC archive has those 32 payloads
  plus the two structural entries.

The production function is `PackageReader::load_parts_eager`, not a separate
tool-local loader (`crates/litchi-opc/src/pkgreader.rs:1279-1309`). Its route is
materially split:

1. `from_phys_reader_with_session` reads and parses the content-types member,
   then loads package-level `_rels/.rels`, before calling `load_parts_eager`
   (`crates/litchi-opc/src/pkgreader.rs:1036-1072`). Those structural reads are
   outside the payload callback.
2. `classify_part_members` skips the content-types member, classifies
   relationship members separately, and appends only non-relationship typed
   members to `typed_parts` (`crates/litchi-opc/src/pkgreader.rs:1538-1606`).
   `_rels/.rels` therefore cannot be one of the selected payload names; a
   per-part `.rels` member would likewise be structural. The fixture has no
   per-part `.rels` members.
3. `load_parts_eager` builds `member_names` from `typed_parts` and passes that
   list to the session callback (`crates/litchi-opc/src/pkgreader.rs:1324-1351`).
   For this fixture the callback receives the 32 `custom/member*.bin` names,
   not `[Content_Types].xml` or `_rels/.rels`. The callback's all-name task-floor
   gate is in `OpenSession::read_many` (`crates/litchi-opc/src/execution.rs:141-227`).

Consequently, the 64 KiB task floor is evaluated over the 32 payload names.
With 32 tasks and a 16 MiB in-flight limit, this corpus fits one batch for all
three shapes. At floor 0, the aggregate 64 KiB threshold is met even by the
small shape (32 × 4 KiB = 128 KiB), subject to requested width greater than one.
At floor 64 KiB, small and mixed are intentionally serial because the
all-members gate sees a 4 KiB member; large remains eligible. A flat mixed curve
at that floor is therefore an admission-policy result, not evidence that the
parallel scheduler failed.

## CFB 4 KiB classification

`build_cfb` emits one named stream per payload (`tools/perf-execution/src/main.rs:425-433`).
The CFB header cutoff is exactly 4096 bytes (`crates/litchi-cfb/src/file.rs:1009-1017`),
but a stream is MiniFAT only when `size < cutoff` (`:1963-1969`). The writer
also rejects `>= 4096` from MiniFAT allocation (`crates/litchi-cfb/src/writer/minifat.rs:1-5,74-82,137-147`).

Therefore a 4096-byte fixture stream is a normal FAT stream, not a MiniFAT
stream. All small-shape streams, and the final mixed-shape stream, are normal
FAT. `SharedOleBulkRead` only forces the shared root MiniFAT cache when a batch
contains multiple `is_minifat` requests (`crates/litchi-cfb/src/shared_bulk.rs:149-185,201-213`),
so this corpus does not exercise that MiniFAT convergence path. At floor 64 KiB
the CFB batch has the same all-members eligibility behavior as OPC: small and
mixed serialize the 32-request batch; large can parallelize.

## Fresh, primed, and route scope

The frozen plan has 120 cases: 40 per route, widths 1/2/4/8/32, floors 0 and
65536, and fresh/primed states. Primed cases are the large shape only, so the
state comparison is made where the 64 KiB floor can admit parallel work.

`fresh` means a new sample-owned session/package with a warm immutable in-memory
source. It is not physical page-cache cold. The route implementations make the
scope concrete:

* OPC creates a new `OpenSession` and times the first `from_bytes` open
  (`tools/perf-execution/src/main.rs:895-931`). Priming performs one verified
  open on the same session, drops that package, then times a second open. The
  explicit ZIP session reads bypass the lazy payload cache; priming primarily
  reuses an eligible private pool and retained worker admission.
* CFB prepares `SharedOleFile` and `SharedOleBulkRead` before the clock, then
  times `read_streams` (`tools/perf-execution/src/main.rs:933-988`). Priming uses
  the same session and resets source counters immediately before the timed read.
  With this all-FAT corpus,
  it primarily tests retained local scheduling state, not MiniFAT cache reuse.
* Parts prepares `SourceBackedPackage` before the clock and times
  `read_parts_ordered` (`tools/perf-execution/src/main.rs:990-1041`). Priming
  loads and verifies once, drops the returned `PartBatch`, then times the same
  package. This is the explicit
  cache-hit control: source-backed payload entries remain retained, so the
  timed observer should show cache hits rather than fresh payload reads.

The packet README records the same boundary (`tools/perf-execution/README.md:26-53`;
`docs/performance/results/change-0786/README.md:3-17`). CFB and Parts observer
metrics cover only the timed operation after reset. OPC uses borrowed `from_bytes`,
so external `ReadAt` metrics are not applicable. The timed closure is the route
operation only: `measure` reads process CPU time before the wall clock and after
the operation (`tools/perf-execution/src/main.rs:675-687`), so its CPU interval
can be slightly wider than wall time. Route byte verification, observer
collection, and drops happen after the clock (`tools/perf-execution/src/main.rs:895-1041`).
Warmup reports are discarded; corpus construction/hashing and JSON output are
outside the recorded per-sample route timing. `run`
constructs the report only after every sample returns successfully, and
`write_report` uses create-new output and removes a file if its write fails
(`tools/perf-execution/src/main.rs:1103-1165`); report creation therefore
supplies a post-run receipt rather than a measured route cost.

## Finite budgets and receipt limits

The standalone context creates a fresh root budget per sample. Root memory,
input, output, objects, depth, and work are finite, and
`Limits::with_execution_io` sets Workers, IoConcurrency, and cumulative
CpuTasks to the requested width, requested width, and one million
(`tools/perf-execution/src/main.rs:689-710`). The route policy separately sets
maximum tasks to 32, maximum in-flight bytes to 16 MiB, aggregate threshold to
64 KiB, and the caller's task floor.

Workers and I/O are outstanding reservations. Private OPC/CFB worker permits
can remain attached to an eligible session until that session is dropped;
`CpuTasks` is cumulative and is not released. A primed sample's preload is
outside the wall interval but still consumes the same sample budget, so the
`before_operation` CPU value must be retained and interpreted as part of the
state. `PartBatch` retains its managed output memory/object reservations until
the batch is dropped (`crates/litchi-opc/src/source_backed/batch.rs:27-39`),
while a primed source-backed package retains payload-cache reservations until
the package is dropped.

The current report proves final Workers/IoConcurrency release and the CPU-task
ceiling (`tools/perf-execution/src/main.rs:713-775`). Its
`ResourceSnapshot` contains only those three dimensions
(`tools/perf-execution/src/main.rs:176-205,713-719`); `ResourceLimits` also
lists `output_bytes`, but no output usage is snapshotted. The report therefore
does not independently receipt peak or final Memory, InputBytes, OutputBytes,
Objects, Depth, Work, or cache bytes. These are configured finite ceilings;
where a route charges a dimension it can reject over-limit work, but those
checks are not measurements that can be inferred from the three execution
counters. A release claim for them would need a richer snapshot or an explicit
diagnostic field.

The resulting claim is a useful finite-budget low-level scaling baseline with
exact ordered payload hashes and route-specific source observations. It remains
warm in-memory, format-low-level, and descriptive; requested-width speedup or
Amdahl fits must not be presented as causal production speedup.
