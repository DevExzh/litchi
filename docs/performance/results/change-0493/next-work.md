# Next work after the managed OPC read-ahead batch

Audit date: 2026-09-10.  This is a bounded, read-only review of the current
non-iWork goal and the executable harness seams.  It records the next measured
end-to-end slices; it is not a 0493 acceptance report and it contains no new
benchmark result.

## Evidence boundary

0492 accepted a private exact-versus-4 KiB adapter comparison over the pinned
synthetic DOCX.  The accepted result is useful for request-amplification
priority, but it is not a production construction, filesystem-cold, borrowed,
native-producer, or concurrent result (`docs/performance/results/change-0492/README.md`,
`results-review.md`).

0493 moves the read-ahead state into `litchi-opc`, forwards an explicit policy
through the DOCX source-backed expert API, and exercises publication transition
guards and managed accounting.  Its timed route is still one fresh managed
DOCX full-text read over an in-memory synthetic transport.  The batch itself
records that it does not establish cold-filesystem, borrowed-source,
multicore-scaling, native-producer, or comprehensive CRUD completion
(`README.md:63-64`, `methods.md:1-7`).  At this audit point the final source
stable capture and strict physical-range validator are still listed as gates
(`measurement-review.md:124-129`); passing those gates would not enlarge the
scope above.

Already usable seams are:

* `tools/perf-baseline/src/lib.rs:46866-46977`,
  `publish_docx_source_edit` and `run_docx_source_backed_one_edit_save`, which
  open a source-backed DOCX, replace one existing paragraph, commit, publish to
  a bounded sequential sink, and check output, replay, inverse, stale-target,
  digest, and ordinary-payload-read oracles.  Its measured source is an
  in-memory `InstrumentedSource`; it is not a provider or cold matrix.
* `tools/perf-baseline/src/lib.rs:14923-14968`,
  `build_docx_source_edit_corpus`, which provides a deterministic 20-member,
  media-rich DOCX (16,793,036 archive bytes and eight 2 MiB media members).
* `tools/perf-baseline/src/docx_provider_lifecycle.rs:592-645,777-842`, which
  already separates provider setup from the timed lifecycle and records source
  version, text, cache, allocation, and source-read evidence for owned bytes,
  `FileSource`, instrumented bytes, and bounded range transport.
* `tools/perf-baseline/src/filesystem.rs:1180-1188,2437-2651` and
  `tools/perf-baseline/src/cold_verified.rs:446-588`, which provide a fresh
  child, aligned-file preparation, strict `fincore` residency checks, and a
  positive process `read_bytes` delta.  Current prepared query controls are
  deliberately ineligible; only a lifecycle whose mandatory source work is in
  the timer can use this proof (`filesystem.rs:7186-7194`).
* `crates/litchi-core/src/execution.rs:32-133`,
  `tools/perf-baseline/src/parallel_metrics.rs`, and the existing
  `OpcSourceConcurrentSamePart`/`run_opc_open_session_scaling` paths provide
  bounded worker, lock, in-flight, speedup, and Amdahl accounting.  0493 uses
  one worker and therefore does not establish scaling.

The committed `docs/performance/CRUD_COVERAGE.md` index predates 0491-0493 in
several rows.  The current goal/checklist and the source/harness paths above
are the authority for this review; no global report is updated here.

## Ranked next slices

### 1. Opened DOCX one-edit/save across provider and cold boundaries

This is the highest-impact next measurement because it closes the gap between
the already implemented source-backed edit route and the goal's required
opened-document edit/save plus source matrix (`docs/GOAL.md:272-316`).  It also
tests the checklist distinction between fresh authoring and editing an opened
document (`docs/CRUD_Scenario_Checklist.md:9-12,46-58`).  Read-ahead evidence
currently covers a read-only full-text lifecycle, not the publication path.

The first implementation contract is:

1. Parameterize the existing `run_docx_source_backed_one_edit_save` runner with
   a provider factory and retain `publish_docx_source_edit` as the exact
   control.  Start with the existing medium media-rich corpus and one existing
   paragraph replacement.  Do not create a second DOCX edit implementation.
2. Run the bounded provider/storage cells below.  The first release-sized
   capture is one corpus, one edit, and 30 retained samples per cell with the
   same warmup and order-reversal discipline as 0493.

   | Cell | Scope |
   |---|---|
   | owned exact | `OwnedSource`, unmanaged `Package::from_read_at`, sequential `CountingSink` |
   | instrumented exact | the same lifecycle with physical range/call tracing |
   | file warm | `FileSource` through the existing filesystem child, with setup outside the timer |
   | file cold-verified | aligned regular file, fresh child, `cold_verified::prepare` immediately before the timed lifecycle |
   | short read | caller `ReadAt` that returns bounded short ranges, max range 4 KiB or 64 KiB |
   | delayed range | bounded range source, max 64 KiB, zero delay and the existing 1 ms plus 100 MiB/s minimum-service arm |
   | borrowed | an explicit ineligible/API row; do not turn `&[u8]` into `Vec<u8>` and label it borrowed |

   Sequential publication is the first measured output scope.  Atomic save is
   a separate follow-up cell using the existing temporary-file, flush/fsync,
   and rename contract; its rename and durability time must not be combined
   with the sequential-sink number.  It should initially cover only warm and
   cold-verified `FileSource`, and must seed and verify the destination so an
   injected failure proves that the old destination remains unchanged.

3. For cold-verified rows, prepare the source and all expected values before
   the final cache drop.  `cold_verified::prepare` hashes the aligned source,
   syncs it, advises `DONTNEED`, and checks `fincore`; that setup is outside the
   timer.  Do not run a post-eviction source fingerprint, full source hash, or
   expected-output rebuild before timing, since it can warm the pages.  Open
   the `FileSource` in the child and put mandatory package indexing, edit,
   commit, sequential publication, and package/document drops inside the
   clock.  `cold_verified::complete` must retain the before/after process I/O
   snapshots and reject a zero `read_bytes` delta.  A row with any setup read
   after the residency probe is ineligible rather than “cold”.
4. Preserve separate before/after oracles for every row: source version and
   source fingerprint, source logical calls/bytes, physical offset/request/
   returned ranges, source-read diagnostics, cache successful/failure counts,
   managed `InputBytes`/`Memory`/`Objects` after drop, output bytes and SHA,
   and semantic reopen/readback.  The edit oracle must prove exactly one
   changed paragraph, all eight media payload hashes unchanged, one commit
   operation, forward patch replay, inverse restoration, and stale/foreign
   refusal.  Atomic rows additionally require the destination preimage on sink
   failure and after a cancellation or budget refusal.

The production-path decision is explicit.  For the unmanaged exact provider
and cold matrix, no production adapter change is needed: the existing
`litchi-docx::source_backed::Package::from_read_at` path and publication method
are sufficient (`crates/litchi-docx/src/source_backed.rs:356-424,875-881,1224-1260`).
The harness needs only provider/cold plumbing and stronger before/after trace
records.  A managed read-ahead edit/save claim cannot be obtained by a harness
adapter alone.  The current source path deliberately rejects it in
`main_document_snapshot` when `cache_diagnostics().budget_managed` is true
(`crates/litchi-docx/src/source_backed.rs:1577-1605`), because
`PartData::into_arc` would escape the managed reservation.  Enabling that arm
requires one separately reviewed production slice for a bounded,
budget-accounted owned edit snapshot or source-bound overlay, with the exact
read transition before snapshot/publication and release on drop.  Removing the
typed refusal or detaching an unreserved `Arc<Vec<u8>>` is not an acceptable
measurement shortcut.  Until that slice exists, report managed edit as a
typed refusal and measure the unmanaged provider/cold route above.

### 2. Bounded concurrent source-backed lifecycle scaling

The goal requires explicit bounded parallelism, worker scaling, and a
concurrent-composition result (`docs/GOAL.md:570-603`; checklist
`CRUD_Scenario_Checklist.md:58,718-735`).  0493's one worker, one read-ahead
window proves transition safety, not throughput scaling.

First add an end-to-end DOCX case that gives each task its own immutable
source-backed package, edit transaction, and sink.  Reuse
`ExecutionLimits`, `default_execution_workers` (1/2/4/8),
`InstrumentedSource.max_in_flight_reads`, `parallel_metrics`, and the existing
OPC open-session scaling runner.  Begin with exact reads and the medium corpus;
then repeat only the delayed and short provider arms if worker 1 shows enough
source latency to expose overlap.  A shared mutable package is a separate
correctness case: independent edits must join deterministically or return a
structured conflict, and it must not be folded into the independent-session
throughput number.

Each worker count needs p50/p95/p99, operations and source bytes per second,
source max-in-flight, lock wait/acquisition data, input/memory/object peaks,
semantic text and output digests, cancellation/failure behavior, speedup,
efficiency, and an explicit serial-fraction/Amdahl calculation.  The expected
first slice is harness-only for independent sessions.  A shared-package
publication result would require a production API/locking decision after the
correctness proof.

### 3. Independent-producer DOCX/PPTX edit/save

The goal requires multiple real-world producers and the checklist requires an
identified producer/version with save/reopen evidence (`docs/GOAL.md:232-244`,
`docs/CRUD_Scenario_Checklist.md:64-69`).  The 0493 corpus is synthetic and
explicitly cannot support this claim.  `tools/perf-baseline/src/pptx_native_image.rs`
currently covers selected-image access/resource observations for Apache POI
PPTX fixtures, and `test-data/office-interop/libreoffice-resaved/` provides
LibreOffice-resaved examples; neither is a certified native Word/PowerPoint
opened edit/save round trip.

The bounded first slice is one stable DOCX fixture and one stable PPTX fixture
from an identified producer, each with one existing semantic edit.  Feed the
same `run_docx_source_backed_one_edit_save` and
`run_pptx_source_backed_one_edit_save` seams after adding fixture-specific
selectors and independently recorded expected semantics.  Reopen the output
with the producer where available, and retain package-level checks for every
untouched relationship, media member, extension, and source/version fence.
The independent producer's output or a separately generated semantic oracle
must establish the expected result; a Litchi-generated expected ZIP alone does
not establish producer evidence.  Only after this correctness route is stable
should it reuse the provider/cold and atomic cells from priority 1.

## Unmet architectural and coverage items kept visible

* Genuine borrowed lifetime remains a production API gap.  `SliceSource<'a>`
  exists (`crates/litchi-core/src/source.rs:187-227`), but the retained
  `SourceBackedPackage` uses an effectively static `Arc<dyn ReadAt>`.  The
  scoped `BorrowedSourceBackedPackage<'a>`/`BorrowedPackage<'a>` contract,
  catalog parity, pointer/range identity, typed Deflate refusal/materialization,
  and compile-fail lifetime tests are specified in
  `docs/performance/results/change-0491/borrowed-source-next.md:8-43,178-220`.
  This cannot be closed by copying bytes into an “owned” benchmark arm.
* The broader opened-document checklist remains partial for bulk updates,
  delete/reorder/closure operations, merge/split, and patch composition.  The
  authoritative required rows are `docs/CRUD_Scenario_Checklist.md:46-81`; the
  stale aggregate index is `docs/performance/CRUD_COVERAGE.md:837-861`.
* Native, malformed/limit, security, cross-format, and iWork rows retain their
  existing scope.  No iWork artifact is used to certify these non-iWork gaps.

