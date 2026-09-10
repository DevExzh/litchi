# CRUD coverage review: next bounded non-iWork work

This pre-implementation review is a read-only audit of `docs/GOAL.md`, the CRUD index, and the
`tools/perf-baseline` selectors. No indexed status is promoted here.

The current validator is deliberately two-stage. Running
`python3 tools/validate_crud_coverage_index.py` succeeds with
`15 categories, 33 mapped selectors (contract-only; no run timing report
supplied)`. The index's `measured` rows require
`target/perf/container-baseline.json` and at least 15 samples, but that report
is absent in this checkout; passing the validator without `--report` therefore
does not establish timing evidence. This is the current index's full-run
validation gap, even though individual historical measurements may exist in
retained change bundles.

## 1. Managed DOCX atomic paragraph batch

This is the strongest next implementation. It closes the explicit managed
refusal and turns existing coalescing evidence into an opened, source-backed
CRUD scenario.

- Production seam: `crates/litchi-docx/src/document/transaction.rs`.
  `replace_body_paragraph_texts` currently calls `ensure_unmanaged` and
  refuses managed transactions. The managed scalar path is
  `replace_managed_paragraph_text`, which reconstructs from the immutable base
  and performs candidate readback for each call. Add one bounded managed batch
  planner/publication path that keeps the same source proof, budget admission,
  operation ledger, semantic readback, stale/refusal, and exact no-op rules,
  but builds one base-relative candidate for the selected paragraphs.
- Harness seam: the planned isolated
  `crates/litchi-docx/examples/managed_paragraph_batch_perf.rs` plus its
  `source_backed_managed_paragraph_batch` focused test. It measures full
  open/edit/commit/sequential-publication lifecycle and edit phase on fixed
  128/512-paragraph corpora with K=1/8/32 selected paragraphs, using owned and
  warm-file providers. This is deliberately fixed-count evidence, not a 1%
  claim. The existing `docx-managed-edit` selector in
  `tools/perf-baseline/src/docx_managed_edit.rs` remains the one-paragraph
  managed control; `docx_semantic_one_percent_edit_save` in
  `tools/perf-baseline/src/lib.rs` is an unmanaged batch guardrail. Keep
  source counters, finite-budget release, reopen, raw/media preservation,
  patch/replay and stale/foreign refusal as gates.
- Scope of 0500 evidence: the isolated example closes managed batch
  capability and gives a scalar-versus-batch lifecycle comparison, but it is
  generated warm-source evidence and should not promote a CRUD-index row or
  claim the full 1%/real-producer/cold-source requirement.
- Evidence: change 0012 already measured the same direct-body coalescing
  shape (large 100-edit/save p50 -94.99%, with the scalar one-edit guardrail
  neutral), while 0495 measured only one managed paragraph and explicitly
  leaves the managed batch helper refused. Thus the implementation is bounded
  and the likely bottleneck is known, but there is no managed-before/after or
  managed scalar-vs-batch lifecycle result.
- Missing required coverage: opened source-backed bulk/1% update under finite
  ownership budgets, plus a real timing comparison. Native producer, cold
  filesystem, high-latency range, unknown-extension, and broad structural
  paragraph coverage remain separate follow-ups.

## 2. PPTX source-backed cross-slide dependency closure

This is the strongest non-DOCX semantic alternative and avoids another OPC
worker/cache experiment.

- Production seam: `crates/litchi-pptx/src/presentation/source_cross_copy.rs`,
  `SourceBackedPresentationEditor::plan_cross_slide_copy` and
  `publish_cross_slide_copy_to_stream`. The current bounded closure accepts one
  slide layout plus direct embedded image/chart leaf parts, and explicitly
  refuses notes, comments, diagrams, tables, external/shared-owner edges and
  richer dependency graphs. A bounded next step is one real producer-backed
  notes relationship or one chart dependency closure with deterministic
  relationship/member collision remapping, preserving fail-closed refusal for
  unsupported topology. The owned counterpart remains in
  `crates/litchi-pptx/src/opened/cross_copy_plan.rs`.
- Harness seam: `tools/perf-baseline/src/lib.rs` functions
  `build_pptx_source_backed_cross_copy_corpus`,
  `run_pptx_source_backed_cross_copy_lifecycle`, and the selectors
  `pptx_source_backed_cross_copy_plain_lifecycle` /
  `pptx_source_backed_cross_copy_media_rich_lifecycle`. They already time
  public source/destination `from_read_at`, planning, and sequential
  publication, and gate semantic reopen, topology, dependency boundary,
  collisions, media payloads, source freshness and raw preservation.
- Evidence bottleneck: the current corpus is synthetic matched plain/media
  (eight 2 MiB image leaves); all cross-copy rows in the CRUD index are
  correctness-only. There is no checked real-producer dependency corpus and no
  release ABBA/full-run report, so the existing lifecycle vectors cannot claim
  a performance result.
- Missing required coverage: cross-document copy of real charts/media/themes/
  notes and collision reconciliation, with source-backed lifecycle timing and
  independent producer reopen. This is a larger semantic closure than the
  DOCX batch and should follow it unless a producer fixture is already ready.

The native RTF and ODF batch selectors remain useful evidence candidates, but
their current gaps are chiefly promotion/producer breadth: the RTF logical
tail is generated plain-only and the ODP text-box batch is already a bounded
owned implementation. Neither is a stronger immediate implementation target
than the managed DOCX refusal.
