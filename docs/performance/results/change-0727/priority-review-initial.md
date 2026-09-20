# 0727 non-iWork priority review

This is a read-only queue review for the next work-elimination measurement. It
uses the optimization order and preservation rules in [GOAL](../../GOAL.md),
the current [hotspot inventory](../../HOTSPOTS.md), and the remaining CRUD
coverage in [CRUD_COVERAGE](../../CRUD_COVERAGE.md). No new timing is taken
here. The 0723–0726 XLS checkpoint and empty-slot variants remain rejected;
0727's native tail replication is diagnostic and does not reopen that queue.

The order below favors a measured opportunity that can remove work on an
existing non-iWork CRUD route, has a bounded first experiment, and fills a
current coverage gap. The figures are evidence from the linked records, not
new performance claims.

## Ranked queue

### 1. Admit stable-index shared strings in the XLSX value-only editor

This is the highest-impact qualification because the reduced readback path is
already implemented but has almost no real-producer reach. Change 0602 measured
the complete candidate parse that the path can remove at **32.9–85.3% of the
commit instruction cost**, and reduced readback at **15.4–43.5% of plan plus
commit p50** on derived producer-shaped sheets. Its real-producer population
was 0/95. Change 0657 widened the surrounding dependency rule and publication
path: 10 packages and 15 worksheet parts were admitted, 9 packages completed
publication, and the remaining largest refusal is the shared-string
relationship. The deterministic structural census in 0657 has 66 shared-string
refusals (62 shared-string-only and 12 combined with pivot-cache refusal) among
the 95-package corpus. This directly extends the partial real-producer XLSX
one-edit/save row in the CRUD matrix.

The source entrypoints are the existing state and refusal seams:

* `crates/litchi-xlsx/src/cell_values/snapshot.rs:1000` keeps the current
  stable-index edit guard; `:1515` defines `SharedStringsState`, and
  `:1937-2005` captures and validates the workbook relationship and part.
* `crates/litchi-xlsx/src/workbook/source.rs:1266-1298` already streams
  selected shared-string dependencies under limits; `:1335-1353` is the lazy
  full-table path. `crates/litchi-xlsx/src/raw/strings.rs` owns the table
  parser.
* The worksheet value parse and staging paths are in
  `crates/litchi-xlsx/src/raw/worksheet/semantic.rs` and
  `crates/litchi-xlsx/src/cell_values/source.rs`.

The first bounded experiment should admit only a source-backed workbook with
exactly one internal `SHARED_STRINGS` or Strict shared-string relationship,
the existing part topology and read limits, and a target cell whose edit does
not add, remove, or renumber a table entry. Keep the shared-string part
byte-identical; continue refusing external relationships, pivot caches,
tables/query tables, metadata, MCE and other value-dependent constructs under
0657's dependency rule. Compare a bounded selected-index lookup or one-time
table materialization against the current refusal on real producer files and a
large-table witness such as the 66,935-entry case from 0602. The gates are
exact edited-cell and untouched-member preservation, index identity,
typed-refusal/error-order parity, cancellation and source fences, and 0602's
F3 price: table materialization must not cost more than the readback work it
unlocks. Do not widen `stored_entry_is_supported` alone; 0602 showed that guard
was unreachable before 0657's admission changes.

Evidence: [0602 design and price](../../0602-xlsx-real-producer-admission-design.md),
[0657 admission/publication census](../../0657-xlsx-value-editor-d4-admission.md),
and the [CRUD matrix](../../CRUD_COVERAGE.md).

### 2. Qualify a length-changing OLE2 copy-through writer for DOC/PPT

The current source-backed CFB overlay and 0663 sector-layout policy cover
same-length overlays or a from-scratch/layout-aware rebuild. The open gap is a
length-changing save that preserves unaffected streams. Change 0617's native
baseline makes this a format-selective opportunity: `FloatingPictures.doc` has
**81.4% unchanged stream bytes**, the container rebuild is **13.25% of commit
cycles**, and work proportional to untouched streams is **12.69%**. The smaller
`NoHeadFoot.doc` case is **29.99%** container and **26.04%** untouched-stream
work. `45543.ppt` has 16.9% unchanged bytes and a 7.26% container share. XLS is
explicitly a poor target: 4.8% unchanged bytes and a 1.07% container share,
inside that measurement's native floor. The rebuild also peaks at 5.2–5.9× the
artifact size, so the work-elimination case includes retention pressure.

The concrete source seams are:

* The existing from-scratch writer is
  `crates/litchi-cfb/src/writer/core.rs:291` and `:1122`; the sequential writer
  is in `writer/sequential.rs:614` and `:757`.
* The common length-changing route enters at
  `crates/litchi-ole-common/src/object/editor.rs:80`, replaces streams at
  `:315`/`:372`, and rebuilds in `:657`. `object/codec.rs:453-504` contains
  `render_with_layout` and the already-landed **same-length**
  `render_copy_through`; that function is not the proposed length-changing
  path.
* The reusable validated primitives are
  `crates/litchi-cfb/src/overlay.rs:569` (`ValidatedOverlayPlan`), `:583`
  (`ComposedOverlaySource`), `:733` (same-length planner), and `:1118`
  (`finish_overlay_plan_with_owner`), with the source-backed owner in
  `crates/litchi-ole-common/src/source_backed_overlay.rs:22`.

The first action is a small ADR/source qualification for the physical-sector
policy, followed by a DOC/PPT-only pilot. The candidate must be a separate
fourth publication path over a validated `SharedOleFile`: existing stream
replacement only; no create/delete/move/rename or storage topology change; a
bounded replacement/output/readback budget; and initial refusal for DIFAT,
signed, encrypted or DRM-marked sources. It should compose unchanged source
bytes with changed stream spans and an appended sector tail, copy directory
records instead of rebuilding them, reopen before publication, and verify every
untouched stream byte-for-byte. Existing fingerprint brackets, source version
checks, FAT/MiniFAT/chain validation, typed errors and atomic publication stay
in force. 0617 recommends reusing only sectors released by this operation,
with append fallback, but that physical placement decision needs the ADR
clarification recorded there before production code chooses it. The pilot must
report DOC/PPT native work and peak live bytes separately; it must not promote
the falsified XLS case into this route.

Evidence: [0617 design and native baseline](../../0617-cfb-copy-through-writer-design.md),
[0663 sector-layout policy](../../0663-cfb-sector-layout-policy.md),
[GOAL's CFB copy-through item](../../GOAL.md), and the [OLE2 CRUD row](../../CRUD_COVERAGE.md).

### 3. Qualify a bounded replay proof for the remaining PPTX cross-slide plan work

Change 0656 removed the second candidate serialization/deflate under a bounded
retained-archive limit. It deliberately left the in-memory candidate graph and
the complete replan intact. On the media-rich lifecycle, planning remains
about **297–301 ms p50** while apply is about **89–94 ms p50**; the accepted
archive reuse removes roughly 23–27% of whole native lifecycle time. The next
measured term is therefore candidate construction and replan work, rather than
another archive-retention attempt. CRUD coverage already has bounded PPTX copy
timing/correctness, while broad dependency closure remains partial, so this
candidate has narrower reach than the first two.

The source path is explicit:

* `crates/litchi-pptx/src/opened/cross_copy_plan.rs:565` plans and retains a
  bounded archive; `:596` applies the plan after semantic and physical revision
  checks.
* `:830` validates the application candidate, `:907` performs the complete
  `prepare_cross_slide_copy_for_slides` refusal/closure work, and `:1157`
  rebuilds the candidate graph and reopens the serialized package.
* `crates/litchi-pptx/src/opened/patch.rs:848` is the final
  `validate_candidate` proof; `cross_copy_plan.rs:2151` owns the physical
  package fingerprint and `:2184` its snapshot cache.

The next measurement should first split `build_candidate` into graph/closure,
serialization and reopen/capture costs, then test whether a compact,
deterministic replay token can remove only graph reconstruction on the clean
owned-source route. It may use a retained archive only when the existing
`max_retained_candidate_bytes` bound admits it and all source/destination
semantic and physical revisions still match. It must retain the 0646/0656
constraint that a plan does not retain a `Snapshot`, and it must not bypass
signature, macro, unknown-member, MCE, protection, collision, limit,
`validate_before`, `validate_after`, target-revision or typed-refusal checks.
Dirty/foreign ingress, missing retention, changed limits, or any proof mismatch
must fall back to today's full replan. If those checks cannot be preserved with
a replay token, close this candidate rather than trusting the archive or adding
a hidden cache; a durable patch or public snapshot representation is out of
scope.

Evidence: [0656 remaining planning cost](../../0656-pptx-cross-copy-candidate-budget.md),
[0646 proof boundary](../../0646-pptx-cross-copy-candidate-retention-design.md),
and the [PPTX copy-closure CRUD row](../../CRUD_COVERAGE.md).

## Explicitly out of this queue

The 0723–0726 XLS checkpoint/empty-slot variants are rejected and remain
diagnostic only. The generic XML namespace rewrite was withdrawn in 0588 and
the authorized writer/fragment rewrite landed in 0653; XML-2 selected-cell
early stopping landed in 0658. PPTX revision memoization (0655), candidate
archive retention (0656), and the CFB Reuse/Rewrite sector policy (0663) are
already implemented. The MCE borrowed-owner experiment has its own retained
tail/refusal evidence in 0702, so it is not a fourth recommendation here.

The small `Namespaces::with_local` empty-vector clone at
`crates/litchi-ooxml-common/src/mce/codec.rs:337` remains a reserve experiment
after these candidates: 0695/0697 attribute it, but the isolated MCE sequence
is only 2.37–2.39 ms and the evidence is narrower than the producer XLSX and
length-changing OLE2 opportunities above.
