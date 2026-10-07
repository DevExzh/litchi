# 0957 commit-drop source trace (v2)

This is a read-only source attribution for the 0956 managed XLSX compression
results. It does not claim a new benchmark result, and it does not change the
repository or run a build/workload.

This is a new v2 note. The v1 report remains unchanged at
`/home/zhuhe/litchi-performance-0957/review/commit-drop-source-trace.md`
(SHA-256 `d986a7ae8a9e83587530f8338e745b2e279a4ab2ef98395562e03772f61c7fc4`).

The production source tree inspected was `HEAD 02449f394529766d76794332204c27c4022931fc`.
The final 0956 benchmark helper used for the timing-boundary citation is
`/home/zhuhe/litchi-performance-0956/source-v1/tools/perf-baseline/src/xlsx_managed_compression_benchmark.rs`,
SHA-256 `bcf634006a1b8069f8c606cb2aa35e4ec560dcddd8c7ff3dfc922301ba582b13`.
The measurements came from the 0956 capture packet
(`/home/zhuhe/litchi-performance-0956/capture-native-v1/children/*/workload-output/report.json`,
144 reports with 200 measured samples each). The reviewed 0956 estimator uses
six alternating ABBA/BAAB process blocks and combines adjacent treatment and
serial quantiles geometrically; samples are not pooled across processes.

## The measured boundary

The 0956 harness separates publication and commit destruction in the final
source-v1 helper:
`/home/zhuhe/litchi-performance-0956/source-v1/tools/perf-baseline/src/xlsx_managed_compression_benchmark.rs`:

- `publish_multi_commit_to_stream` is called at lines 429-431.
- The returned `MultiSnapshot` is dropped at line 431, and the publication
  timer stops at line 432.
- A new timer starts at line 433, `drop(commit)` runs at line 434, and the
  `commit_drop_ns` timer stops at line 435.
- The phase contract is recorded at lines 698-702: `publication_ns` includes
  publication and the returned snapshot drop; `commit_drop_ns` is the
  `MultiCommit` drop.

Therefore the elevated field is the caller-side destruction of the already
published `MultiCommit`. The editor/package is consumed by the publication
call, so its ordinary package destruction is reached before that call returns:
`SourceBackedEditor::publish_multi_commit_to_stream` takes `self` and returns a
cloned target at `crates/litchi-xlsx/src/cell_values/source.rs:600-650`.

The medium one-percent case has five Deflate tasks, the same source bytes
(4,226,429), output bytes (4,226,788), source hash, and output hash at every
width. Only its widths 2 and 4 pass the managed admission preflight with worker
permits. The one-edit case has two tasks but remains serial at all widths; the
dense-sparse one-percent case has five tasks but remains serial at all widths.
The final 0956 report records the reviewed medium one-percent commit-drop p50
ratios as 1.738 for width 1 versus 2 and 1.851 for width 1 versus 4. Across
the twelve processes in each comparison, serial p50 commit drop spans about
167-170 microseconds, versus 291-295 microseconds at width 2 and 306-320
 microseconds at width 4. The corresponding control comparisons remain near
 one. These are the frozen paired process results and reviewed descriptive
 ranges, rather than a post hoc pooled-width estimator. The correlation between
 the width-specific publication route and the higher medium one-percent
 `commit_drop_ns` is therefore supported by the paired evidence, but it does
 not prove that width changes the logical commit graph. The reviewed ranges and
 ratios are recorded in the final repository report at
 `docs/performance/0956-xlsx-managed-compression-native-widths.md`, SHA-256
 `c7f26b20a139b9ad9bc0ad8d6e7acd85a572e23c9b9b36b9b6475939a9d71be0`.

## What `MultiCommit` retains

`MultiCommit` has three fields: `snapshot`, `patch`, and content-free
`diagnostics` (`crates/litchi-xlsx/src/cell_values/patch.rs:217-239`). Its
`MultiPatch` has an owned `before` and an owned `after` snapshot
(`patch.rs:23-26`). In `MultiSourceEdit::commit`, the candidate vector is
frozen into one `MultiSnapshot`; `MultiPatch::new(self.before, after.clone())`
clones that candidate; and `MultiCommit::new(after, patch, ...)` stores the
original candidate as its own `snapshot` (`crates/litchi-xlsx/src/cell_values/source.rs:1146-1254`).
Thus the successful commit retains this ownership graph until
`drop(commit)`:

1. the `MultiCommit::snapshot` candidate snapshot;
2. `MultiCommit::patch.after`, a cloned candidate snapshot; and
3. `MultiCommit::patch.before`, the source snapshot.

`MultiSnapshot` owns a `Box<[Snapshot]>` (`crates/litchi-xlsx/src/cell_values/snapshot.rs:151-202`).
Each `Snapshot` contains a boxed sheet name, an `Arc<Store>`, and a complete
`SourceState` (`snapshot.rs:54-70`). `SourceState` includes `PartState`s,
boxed relationship strings, shared `Arc` relationship/graph arrays, optional
shared-string state, source context/lineage/version, and the detached
replacement-span hint (`snapshot.rs:1466-1482`, `snapshot.rs:1511-1523`).
A `PartState` has a boxed content type and a `SourcePayload`
(`snapshot.rs:1567-1572`).

The payload handles are clone-shared. `SourcePayload::Managed(PartData)`,
`Owned(Arc<Vec<u8>>)` and `OwnedBudgeted { bytes, budget }` are defined at
`crates/litchi-xlsx/src/source_payload.rs:41-52`. `PartData` itself contains an
`Arc<Vec<u8>>` plus managed payload and object reservations
(`crates/litchi-opc/src/source_backed.rs:5108-5160`). A value-only rewrite
creates one `Arc<Vec<u8>>` for the rewritten worksheet and a detached span hint
(`snapshot.rs:772-829`); workbook invalidation similarly creates a shared
`Arc<Vec<u8>>` (`snapshot.rs:861-884`). The three snapshots therefore own
separate top-level boxes and cloned boxed metadata, while many stores, source
payload allocations, replacement payloads, reservations, and hint spans are
released by reference-count decrement. Some of those decrements may be final
frees in `drop(commit)`.

This graph is constructed before publication. There is no worker-width branch
in `MultiSourceEdit::commit`, and the source-backed ZIP writer documents that
scheduling changes when compression runs while preserving each member's bytes,
physical order, central order, and typed errors
(`crates/soapberry-zip/src/preserve.rs:929-943`). The source therefore supports
a fixed logical commit graph across widths; it does not support an assertion
that width 2 or 4 created a larger or different `MultiCommit`.

## What publication owns and releases first

`publish_multi_commit_to_stream` clones the patch target for its returned
`MultiSnapshot`, builds a topology plan, and consumes the editor/package
(`crates/litchi-xlsx/src/cell_values/source.rs:639-650`). The OPC compatibility
wrapper consumes the `SourceBackedPackage` through
`prepare_topology(plan)?.publish_to_stream(writer)`
(`crates/litchi-opc/src/source_backed.rs:10058-10067`). A prepared topology
owns the package, changed overlays, appended regenerated entries, transfer
state, and reservation vectors (`source_backed.rs:2630-2662`); its publisher
moves those values into `publish_prepared_topology` and then into the changed
overlay writer (`source_backed.rs:2773-2818`, `source_backed.rs:2863-2891`).

For a value-only replacement, the changed overlay retains an `Arc<Vec<u8>>`
(`source_backed.rs:2601-2610`, `source_backed.rs:9680-9685`). The ZIP plan
clones that `Arc` into a `RegeneratedEntry::new_shared`
(`source_backed.rs:12097-12108`; `crates/soapberry-zip/src/preserve.rs:144-192`).
Those publication-local `RegeneratedEntry` and prepared-entry handles are
released when publication locals are destroyed; the commit's snapshot still
owns its own clone.

A qualifying parallel wave additionally allocates publication-local scheduling
and result state. `compress_deflate_wave` reserves `entries`, `results`,
`pairs`, worker closures, and task references, then returns a plan-ordered
`CompressedWave` (`crates/soapberry-zip/src/preserve.rs:2335-2455`). Each
Deflate `generated_entry` fully buffers a one-member mini-archive; the finished
buffer is shared by its local and central ranges, and the compressor is
constructed and freed per member (`preserve.rs:2511-2565`). The writer prepares
all regenerated members before emitting the first local record, then drops its
prepared vectors after the forward-only output pass
(`preserve.rs:1006-1077`, `preserve.rs:1120-1193`). These allocations are in
publication, not in the `MultiCommit` value being timed afterward.

## The private pool boundary is not a join boundary

A managed `PartCache` stores an inert `ParallelWriteSession` and optional
`WritePermits` (`crates/litchi-opc/src/source_backed.rs:5448-5462`). Its session
is created without a pool; a local pool is built only when a qualifying wave
runs (`source_backed.rs:5555-5586`). Admission charges one `CpuTasks` unit per
regenerated member at every width. A parallel wave also reserves worker
permits and, without caller-provided workers, retains those permits in the
cache (`source_backed.rs:5640-5693`, `source_backed.rs:5695-5716`). Width one
returns the serial route without permits. Widths 2 and 4 in the medium
one-percent case use the local pool path; the two control groups do not.

`ParallelWriteSession::run_tasks` builds a pinned local Rayon pool and calls
`pool.install` for a parallel wave (`crates/soapberry-zip/src/office/parallel_write.rs:263-309`).
The 0956 context does not attach a caller `ScopedWorkers` facility, so this is
the local-pool variant. The `SourceBackedPackage` owns the cache
(`source_backed.rs:6584-6608`), and the package/session handle is destroyed by
the consuming publication path described above. That is the part correctly
covered by the publication phase.

The pinned dependency is `rayon 1.12.0` / `rayon-core 1.13.0` (the lockfile
entry and checksum are in `Cargo.lock:4008-4028`). In that source,
`ThreadPool::drop` calls `registry.terminate()` only
(`~/.cargo/registry/src/index.crates.io-1949cf8c6b5b557f/rayon-core-1.13.0/src/thread_pool/mod.rs:399-403`).
`Registry::terminate` decrements the termination count and sets each worker's
termination latch; its comment says workers “will gradually terminate, once
any extant work is completed” (`.../rayon-core-1.13.0/src/registry.rs:591-600`).
The implementation does not wait for all OS worker threads to stop. The only
`wait_until_stopped` helper is test-only (`registry.rs:396-403`), and the
worker's `stopped` latch is documented as test-only (`registry.rs:613-627`).

Consequently, the publication interval contains the pool handle's drop and
termination signal, but it does not establish that worker-thread teardown has
finished before the following `commit_drop_ns` interval starts. `pool.install`
does wait for the scheduled compression work to finish before the writer
continues, so this caveat concerns post-wave worker shutdown and thread-local
cleanup, not unfinished Deflate tasks. The 0956 zero-after-drop managed
resource checks establish that litchi budget leases were released; they do not
measure OS thread quiescence or native allocator state.

## Attribution and hypotheses

The following points are supported directly by the source and results:

- The measured interval calls Rust `drop` on a fixed-shape `MultiCommit` after
  publication and returned-snapshot destruction.
- Medium one-percent is the only measured case where the width setting admits a
  local parallel wave. The reviewed paired commit-drop ratios for the two
  non-parallel controls remain near one.
- Publication creates and destroys additional parallel-wave vectors, prepared
  member buffers, compressor state, and worker-pool state before the commit
  timer, while the commit retains the three snapshot graphs described above.
- The pool drop signals termination but does not join worker threads, so worker
  exit can overlap the caller's subsequent commit destruction.

The likely explanations below are hypotheses, not source-proven causes:

1. **Asynchronous worker exit and thread-local cleanup.** Width 2 or 4 can leave
   more Rayon worker state and more worker-local queue/allocator state to unwind
   after the pool's termination signal. That unwind can run concurrently with
   `drop(commit)` and can alter scheduling or allocator contention. Width 4's
   higher p50 than width 2 is consistent with this, but no worker-stop timestamp
   or thread identity was recorded.
2. **Allocator state left by the parallel wave.** Parallel publication allocates
   the `CompressedWave` vectors and one buffered generated archive per Deflate
   member, and frees those objects on the worker/caller paths described above.
   Different allocation and free ordering can leave different thread-cache or
   arena state before the three snapshot graphs are destroyed. This could make
   the same `Arc`, boxed-string, and reservation drops take different time. The
   repository source does not select an allocator or measure allocation events,
   so this remains a timing hypothesis.
3. **Final shared-handle releases.** `drop(commit)` may perform the final
   releases for some rewritten `Arc<Vec<u8>>`, managed `PartData` reservations,
   `Arc<Store>` values, boxed relationship metadata, or replacement-span hints.
   The source proves that these handles are present and clone-shared, but it
   also proves that the logical graph is width-independent. Width can therefore
   change the surrounding allocator/thread state, but this trace does not show
   width changing the number or size of commit-owned handles.
4. **Direct pool/permit destruction is not the measured cause.** The package's
   pool handle and `WritePermits` live in the consumed editor/package and are
   released on the publication path. They explain why the parallel route is
   different, but their synchronous destruction belongs to `publication_ns`.
   Only asynchronous worker termination after the signal can overlap
   `commit_drop_ns`.

The evidence does not justify attributing the extra 125-143 microseconds to a
larger ZIP output, a different `MultiCommit` topology, or a second source
publication. The source and output hashes are identical across widths, and the
benchmark drops the returned target before starting the commit timer.

## A clean follow-up separation

To identify the remaining cause without changing the canonical timing contract,
a diagnostic run would need at least one of the following outside the measured
result: a test-only worker-stop latch/barrier after package destruction, a
caller-provided `ScopedWorkers` route that keeps the wave admission but removes
the private Rayon pool, or allocator/thread event counters separated for
publication and `drop(commit)`. A worker-stop wait would itself belong to a
diagnostic boundary; adding it to `publication_ns` would change the 0956
contract. The current reports contain no such evidence, so this review stops at
the source-supported boundary and the explicitly labeled hypotheses above.

## Input receipts

The following hashes identify the source files read at this HEAD and the pinned
pool implementation. They are included for attribution reproducibility; no
repository files were modified.

| input | SHA-256 |
| --- | --- |
| `crates/litchi-xlsx/src/cell_values/patch.rs` | `762c222c7fc3580a66c4c1f11f9de44d4a43d4203210597cacdfe34fc17e4648` |
| `crates/litchi-xlsx/src/cell_values/snapshot.rs` | `50a4ef8fd9f2e810d2b1315d8a389086134a28901a2b86a98cc8abeb84308726` |
| `crates/litchi-xlsx/src/cell_values/source.rs` | `60d89d8af3847d3df5c96801537e58ec676446fb5d1828b4adbc54d92fc6a462` |
| `crates/litchi-xlsx/src/source_payload.rs` | `20d90a45659711e078c5b376823952d04b90a325e69ce6f3e3797b7ea17fff1f` |
| `crates/litchi-opc/src/source_backed.rs` | `571ec3e2e08d11626b5ce489c5d7de498445c1dd72441692306772c1893c3498` |
| `crates/soapberry-zip/src/office/parallel_write.rs` | `c8228b7ce57a11338a5a3e84480ac23f0d1d7e47760f1245b2050161390954f0` |
| `crates/soapberry-zip/src/preserve.rs` | `21e679a6c0cdc027859d13c87a24564fba3ba59f7ecfc651b971930b1fa063ae` |
| `Cargo.lock` | `d80aa535dcf880a85fa31e92fe2e5ce78b91d93a55953ae734e513aed83e0061` |
| 0956 representative medium-one-percent report | `169af13ccfa6b30e0616dbabe9979e57b70348a7c4f614d30483a794b7609afe` |
| final 0956 repository report | `c7f26b20a139b9ad9bc0ad8d6e7acd85a572e23c9b9b36b9b6475939a9d71be0` |
| final 0956 benchmark source-v1 | `bcf634006a1b8069f8c606cb2aa35e4ec560dcddd8c7ff3dfc922301ba582b13` |
| pinned `rayon-core-1.13.0/src/thread_pool/mod.rs` | `f071a1bc766cde10ec844cc3b53e53479c3f1ade97ceaff1cc9e3f9913ade9f3` |
| pinned `rayon-core-1.13.0/src/registry.rs` | `61674f4cdf39edb26df75435de6925d54c5f864f9e1dcd0fb114ad2cc553a1d7` |
