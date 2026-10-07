# 0957 allocation design review

Reviewed at 2026-10-07T12:28:02Z. This is a read-only design review; no
source, build, benchmark, or workload was changed or run.

## Evidence and scope

The review is bound to checkout HEAD `02449f394529766d76794332204c27c4022931fc`.
At review start, the only pre-existing worktree changes were
`tools/check_iwork_leaf_public_api.py`, `crates/litchi/examples/bench_word_parse.rs`,
and `docs/UNIFIED_OPS_API_DESIGN.md`; this review did not touch them. By final
handoff, a concurrent 0957 implementation had also modified
`tools/perf-baseline/Cargo.toml` and `tools/perf-baseline/src/lib.rs`. Those
changes are outside this read-only review and were not built or validated here;
the declared `xlsx_managed_compression_alloc` source was not present in the
shared checkout at handoff.

Relevant source identities at review time are:

| File | SHA-256 |
| --- | --- |
| `tools/perf-baseline/src/allocation_metrics.rs` | `0fcec3e5c4972031b23542fd359a38401e25f32a133c77ef4c6e75facf0710d5` |
| `tools/perf-baseline/src/xlsx_managed_compression.rs` | `d2b2de1f103590acbe52bcdecad930d2f2c80d61e515ea3969cf9b2f6a17e148` |
| `tools/perf-baseline/src/xlsx_managed_compression_benchmark.rs` | `bcf634006a1b8069f8c606cb2aa35e4ec560dcddd8c7ff3dfc922301ba582b13` |
| `tools/perf-baseline/src/bin/litchi-perf-baseline-alloc.rs` | `69d26775e4fb17405de0e01fc83a04ce57191c66b093e4765d6b9301824a400a` |
| `tools/perf-baseline/src/bin/support/counting_allocator.rs` | `80977e55c2213eeee2af9a3133295d70cff38505e663a68e4757bf7939c60068` |
| `crates/litchi-xlsx/src/cell_values/source.rs` | `60d89d8af3847d3df5c96801537e58ec676446fb5d1828b4adbc54d92fc6a462` |
| `crates/litchi-opc/src/source_backed/batch.rs` | `e4bb877c1aa54d5cf13617b547ce33a2aff63c6a7cf2a3984d49aa384b85922b` |
| `crates/soapberry-zip/src/office.rs` | `f72bfef6e6a072b3e88c465afb1358e4f77046e37c606ed3902fe3a1621ceac3` |

The worker/ownership hashes are source identity checks for the boundary
references below; they are not proposed production changes.

The governing files read for this design were `docs/GOAL.md`,
`docs/adr/README.md`, ADR 0005 (`8279fd2dc0b074aa104bf0c83df6187d12e18756cd8c8d7f25750b2a997cb0aa`),
ADR 0031 (`ce7e16baa101cfdd8abfc6d606dc0a56fad6d86cc79eba997d338f919a063eb3`),
and ADR 0032 (`33e68f71a2845132ff53aec30422300194b1bc40caa6dd9503848278e342ed70`).
ADR 0005 requires representative allocation/peak-memory evidence without
unsupported RSS or throughput conclusions. ADR 0031 makes worker width and
resource admission explicit and rejects ambient process-global scheduling.
Nothing in this design changes either production execution policy or budget.

The 0956 native report is `docs/performance/0956-xlsx-managed-compression-native-widths.md`.
It says that every iteration creates a fresh finite managed context, package,
and bounded sink; source/context/cache/sink preparation is outside the native
clock; the cold managed session/private pool is inside the process; publication
includes the consuming editor and its package/private-pool teardown; and
commit-drop is a separate phase. The report explicitly makes no allocation
claim. The allocation follow-up must preserve those route and oracle contracts
while making a separate claim about allocator counters only.

## Current implementation facts

`allocation_metrics.rs` is already the appropriate safe observer. `enable()` is
called only by the isolated allocator executable. `Snapshot` carries allocation,
deallocation, reallocation, failed-allocation, allocated/deallocated-byte,
process-live, process-high-water, overflow, and observer-invalid fields
(lines 79–140). A `Region` is non-reentrant and uses one observer mutex to
linearize allocator callbacks and boundaries (lines 295–444, 507–618). Its
`split()` rebases a non-overlapping segment while retaining the complete-region
peak, and `finish_split()` returns both the last segment and the combined sample
(lines 300–379).

The important interpretation is:

* `live_bytes_before/after` are absolute process-wide net live bytes. They
  include allocations made before the phase, including the corpus, report
  buffers, allocator/runtime state, sink, and any context that was deliberately
  pre-created.
* `region_peak_live_bytes` is the observer-ordered high-water live value for
  the phase. It includes successful callbacks from any process thread while the
  region is active, retains an entry allocation even if it is freed later, and
  is not RSS or a unique-heap measurement. The module documents that it can
  perturb allocator scheduling and excludes allocator-internal realloc overlap
  (lines 9–19). Tests cover concurrent peaks and cross-thread totals
  (lines 1098–1153).
* `peak_live_bytes_before/after` are monotonic process high-water context, not
  phase peaks. Do not use their difference as phase allocation.
* `allocated_bytes` and `deallocated_bytes` are traffic. A successful realloc
  increments both traffic totals (new and old sizes), increments
  `reallocation_calls`, and changes net live by the size delta (lines 699–724).
  Therefore `allocation_calls` is not a malloc-only count.
* `Sample::measured` fails closed when arithmetic wraps, observer state is
  invalid, a peak is missing, or the peak is outside the before/after/global
  bounds (lines 222–269). A diagnostic must reject `Unavailable` and `Overflow`
  rather than publish partial vectors.

The isolated allocator wrapper in `counting_allocator.rs` is the only unsafe
boundary. It delegates to `std::alloc::System` and records successful and failed
alloc/realloc/dealloc callbacks (lines 16–58). No new unsafe code or allocator
implementation is needed.

The current `litchi-perf-baseline-alloc` selector has no
`xlsx-managed-compression` arm (lines 14–48); its default path is not this 0956
route. Consequently an allocation run is not available today. A future change
should add one allocator-only selector/child module and a separate allocation
schema. The normal executable must remain uninstrumented, and the native 0956
runner/reducer must never consume allocator reports.

The current 0956 route is explicit in
`tools/perf-baseline/src/xlsx_managed_compression_benchmark.rs`:

* `timed_iteration` creates the `Budget`, cancellation pair, execution context,
  fresh `OwnedSource` clone, and resets the sink before the native clock
  (lines 401–412).
* Ingress is the `SourceBackedEditor` constructor (lines 413–422).
* Edit/commit is selector construction, cell updates, and `MultiCommit::commit`
  (lines 424–428).
* Publication writes the commit and drops the returned snapshot (lines 430–433).
* `MultiCommit` is dropped in its own phase (lines 434–436).
* The native verifier checks changed cells, `OutputBytes`/`CpuTasks`, clean
  resources, exact bytes and hashes, semantic output, and inverse restoration
  (lines 465–534). It then drops the caller's `Budget` at line 535.
* `BoundedSink::new` is called by `run_case` at lines 639–653 and preallocates
  its full capacity at lines 128–147. The sink is therefore outside the current
  native timers, but is a fresh per-iteration owner.

The public XLSX publisher confirms the ownership boundary:
`SourceBackedEditor::publish_multi_commit_to_stream` takes `self` and returns a
`MultiSnapshot` (`crates/litchi-xlsx/src/cell_values/source.rs:599–651`). The
editor, its source/context owner, and publisher-owned package/scheduler state
are consumed in publication. They must not be assigned to `commit_drop` or a
later post-drop phase merely because their deallocation callbacks happen near a
call return.

## Smallest trustworthy diagnostic

Add a child module under the existing benchmark module, reusing its private
`CasePlan`, `BoundedSink`, `build_case_plan`, `stage_timed_commit`, fixed source
and output hashes, and inverse/oracle helpers. Add one dispatch arm to the
allocator binary only. Leave `timed_iteration`, the normal executable, and the
native report schema unchanged. The child should use one process per case/sample
or prove worker quiescence between samples; it must not turn allocator
instrumentation into a native timing comparison.

Use preallocated report vectors and fixed strings before the first allocation
region. No JSON serialization, `format!`, ordinary procfs read, logging, or
dynamic error construction may occur while setup or one of the four native core
phase segments is active. The only in-region observer is the fixed-buffer drain
reader described below, after its zero-allocation qualification. The route should
use exactly the 0956 corpus, update set, output/sink ceiling, finite context,
worker widths, and inverse checks. A first bounded attribution set is medium
one-percent widths 1/2/4; dense-sparse and one-edit can remain controls under
exactly the same route contract.

### Phase boundaries

The least invasive layout has three disjoint allocator regions:

1. **`setup`**: begin before creating the fresh bounded sink, finite
   `managed_context`, cancellation pair, and `OwnedSource` clone; finish just
   before the editor constructor. This captures cold context/source/sink setup
   while leaving the corpus, expected output, inverse patch, and all report
   storage pre-existing. Record the pre-setup process live value in the sample.
2. **`core`**: begin immediately before the editor constructor and retain one
   `Region` through the four requested core phases and an explicitly separate
   worker-drain segment. Call `split()` immediately after editor ingress, after
   `stage_timed_commit`, and after publication (including returned snapshot
   drop). Call `split()` immediately after `drop(commit)` to publish
   `commit_drop`; then drain the private pool while the same region remains
   active. Finally call `finish_split()` to publish `worker_drain` and the
   `core_combined_with_drain` sample. The four native phase samples are never
   merged with the drain sample; the combined sample is labeled as including
   the drain. The publication segment includes the consuming editor and its
   source/context/private scheduler teardown, as the public API requires.
3. **`postdrop`**: run all correctness checks with no allocation region active,
   then begin a fresh region and drop the remaining caller-owned `Budget` and
   per-iteration sink. Finish before pushing/serializing samples. This captures
   the cleanup that the native verifier currently performs at line 535 and the
   sink owner that `run_case` otherwise drops at the end of its loop. If the
   implementation elects to keep the sink outside the diagnostic lifetime,
   call the result `budget_postdrop` and state that sink release is unavailable;
   do not silently call it full teardown.

The core `Region` must be finished before verification. Calling the existing
`verify_iteration` directly would drop the budget at line 535 and prevent an
accurate postdrop phase. The child should use a private, line-for-line
`verify_before_cleanup` helper that performs the existing changed-cell,
budget, sink, exact-output/hash, semantic, source-hash, and inverse gates but
returns before dropping the budget. This preserves the native function
unchanged and avoids a fake zero-duration bridge. A private helper extraction
shared by native and diagnostic code would also be sound, but is a larger
source diff and needs a separate source review.

`postdrop` is not a claim that all process memory vanished. The corpus, case
plan, allocator/runtime state, and preallocated report storage intentionally
remain live. If process-level teardown of a worker pool is deferred, the
postdrop sample is incomplete; the runner must mark it unavailable rather than
assigning delayed callbacks to a prior phase.

Because setup and core are intentionally separate regions, report
`setup_peak`, `core_combined_with_drain_peak`, and a derived checked
`max_setup_core_peak` separately. The derived value is the maximum of two
non-overlapping observer intervals, not a continuous single-region peak. Keep
all raw before/after/peak fields so the pre-existing live baseline is visible.

### Report contract

Use a new allocation-only schema, for example
`litchi.xlsx-managed-compression-allocation.v1`, with these fields per sample:

* route identity: case, shape, workload, workers, source/output/semantic hashes,
  update count, context limits, allocator identity, binary identity, scope, and
  counter revision;
* `setup`, `ingress`, `edit_commit`, `publication`, `commit_drop`,
  `worker_drain`, `core_combined_with_drain`, and `postdrop`
  `allocation_metrics::Sample` values;
* checked `max_setup_core_peak_above_setup_live_before`, plus explicit
  `peak_semantics: observer_ordered_process_live_not_rss`;
* `workers_requested`, expected/admitted `Resource::Workers`, `CpuTasks`, and
  `resources_after_drop`; these are budget/admission evidence, not OS thread
  counts or per-thread allocator attribution;
* `thread_observation_status`, first/final drain counts, poll count, deadline,
  and parser status. Ordinary process observations remain outside measured
  regions; the fixed-buffer drain is explicitly part of `worker_drain` after
  its zero-allocation qualification. Neither form fills a per-thread byte
  breakdown.

Every phase requires `status == measured`, `observer_invalid == false`, and all
numeric fields present. The route must also assert the existing output bytes,
output/semantic/source hashes, sink bounds, changed-cell count, exact inverse,
`OutputBytes`/`CpuTasks`, cancellation/error behavior where applicable, and
zero retained managed resources. Allocation counters supplement those gates;
they do not replace them.

### Cross-thread and counter controls

The global wrapper and observer mutex count callbacks from all process threads
while a region is active. This is enough to include private-pool worker
allocations in the phase totals, but it does not identify the allocating thread
or prove that a permit became a live OS thread. Use the following controls:

* Run the same allocator-only route at worker widths 1, 2, and 4. Report
  `Resource::Workers` admission and `CpuTasks` separately. The 0956 report
  already says permits are admission evidence, not task overlap; retain that
  limitation.
* Retain `allocation_calls`, `deallocation_calls`, `reallocation_calls`,
  `failed_allocation_calls`, allocated/deallocated traffic, live before/after,
  and observer peak for every phase. Do not infer malloc count from
  `allocation_calls`, because reallocations contribute to it.
* Keep the existing allocator unit controls: concurrent thread allocations and
  exact overlapping peaks (`allocation_metrics.rs:954–987` and
  `1098–1153`), realloc growth/shrink and failed allocation accounting
  (`1024–1047`), nested-region rejection, and overflow/reentry fail-closed
  behavior. A future child-level control may use pre-started safe scoped worker
  threads to allocate, reallocate, and deallocate known vectors, but thread
  startup must occur before the measured region unless startup is deliberately
  part of `setup`. No production unsafe code is needed.
* Do not add a thread-id map, global allocator registry, or hidden pool. A
  lock-free harness-only active-worker gauge could document task overlap, but it
  would be extra instrumentation and is not needed for the first allocation
  diagnostic; absence should remain `unavailable` rather than inferred from
  permits.

The child must establish scheduler quiescence before ending `core` or
`postdrop`. The local OPC batch path joins its scoped workers, but this XLSX
route can also reach `soapberry_zip::office::ParallelReadSession`, whose source
documents a lazily created private Rayon pool (`crates/soapberry-zip/src/office.rs:263–303`,
`582–601`). Dropping that pool is not an attribution boundary by itself: a
termination signal can leave worker callbacks after the publisher returns. A
zero `Resource::Workers` value proves only budget permits were released; it does
not prove OS worker threads have exited.

For the Linux host used by 0956, pre-open `/proc/self/status` before the first
region. Read it with a fixed 16 KiB stack buffer and parse the `Threads:` field
without `String`, `Vec`, formatting, or allocation. First qualify this observer
inside an otherwise empty allocator region (100 polls at the one-thread
baseline): require measured status, zero allocation/deallocation/reallocation/
failed-allocation deltas, and unchanged live/peak values. The same observer can
then run as the fifth `worker_drain` core segment: after the four native phase
splits, poll/yield to a bounded deadline until `Threads` returns to the recorded
baseline, and only then call `finish_split()`. Record the first post-core count,
final count, poll count, deadline status, and parser status. Do not poll between
the four native phase boundaries. Procfs observer time is diagnostic boundary
work, not a native timing measurement.

If baseline is not one thread, procfs is unavailable, the zero-allocation
qualification fails, or the drain deadline expires, retain the four native
phase samples but mark `worker_drain`, `core_combined_with_drain`, and any
postdrop attribution unavailable. A fresh allocator process per route sample is
the fallback when a reliable drain cannot be established. Never charge a
callback observed after the boundary to publication or commit-drop merely
because the resource permit was released.

## Interpretation limits and recommendation

The wrapper's observer mutex and callback overhead can affect scheduling; an
allocator target's elapsed time, if emitted, is not comparable to the 0956
native timing ratios. The allocation report should contain no `Instant` phase
performance claim and should not be sent through the native reducer. RSS,
page faults, and context switches remain separate whole-process observations.

This design can answer: “which route phase produced allocator traffic and
observer live peaks under the cold finite managed context, including callbacks
from worker threads?” It cannot answer: “how many bytes each worker allocated,”
“what was unique RSS,” “how many OS threads actually ran,” or “why the native
latency changed.” Those require separate instrumentation and evidence.

The minimal next implementation is therefore an allocator-only child runner,
with setup/core/postdrop regions and existing safe counters, first exercised on
medium one-percent widths 1/2/4. Preserve all 0956 semantic and resource
oracles, run correctness outside allocation regions, retain raw status fields,
and report no native timing or production/adoption conclusion.
