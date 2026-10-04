# 0906 legacy DOC/XLS/PPT coverage audit

This is a finite, read only audit of the legacy DOC, XLS and PPT paths. No
build, test, workload or performance capture was run here. The shared checkout
was at `59f75c7f79026a4d124b61e97c178c7b45de19ab`; it contains unrelated iWork
conflicts and edits, which are outside this audit. The source context binds the
audited perf source to office branch `2b974e3541517ccf032021dc5497243b315fd313`
and records the source hashes in `source-context.json`.

The 0906 lexical registry contains 534 selectors: 41 default selectors, with
14 `doc_*`, 40 `xls_*`, and 29 `ppt_*` names. The registry replay passed with
zero builds, workloads, and worktrees. These are callable selector mappings,
not measurements, corpora, or proof that a CRUD direction is supported. The
perf README makes that distinction explicit. Only the three fresh legacy writer
selectors are in the 41-case default; all legacy semantic, source, facade and
edit selectors are opt-in.

After the source-context snapshot, the shared checkout acquired an uncommitted
15-line perf help/test mapping change in `tools/perf-baseline/src/lib.rs`; its
current hash is `f712da842169399d687eb452c7042deed61682a80efdd40b2c3c178af2c036ea`.
The diff is outside the legacy enum, runners and corpus builders audited here.
The source-context hash (`b62775d903ccb2490c28af61d6c5c95d8ed826afad4f0e7dc3c83bd41b877903`)
is retained below as the immutable capture identity; both hashes are listed so
the current dirty state is not mistaken for the captured source.

## Exact callable legacy cases

The case enum, names, CLI mapping and runners are in
`tools/perf-baseline/src/lib.rs:1104-1640,1686-2061,2514-2638,25213-25308`.
The current callable operations are:

- **DOC (14):** fresh authoring (`doc_fresh_write_to`); synthetic semantic open,
  paragraph list/count, one paragraph, positional paragraph, full text, exact
  no-op save and one paragraph replacement
  (`doc_semantic_open`, `doc_semantic_list_paragraphs`,
  `doc_semantic_paragraph_count`, `doc_semantic_one_paragraph`,
  `doc_semantic_one_paragraph_at`, `doc_semantic_full_text`,
  `doc_semantic_noop_edit_save`, `doc_semantic_one_edit_save`); body snapshot
  paragraph listing (`doc_body_snapshot_list_paragraphs`); public phase
  attribution (`doc_owner_public_phases`); and named real-file facade open,
  full text and one paragraph (`doc_facade_file_open`,
  `doc_facade_file_full_text`, `doc_facade_file_one_paragraph`).

- **XLS (40):** synthetic open, worksheet list, one cell, full cell scan,
  exact no-op and one numeric edit (`xls_semantic_open`,
  `xls_semantic_list_worksheets`, `xls_semantic_one_cell`,
  `xls_semantic_full_cell_scan`, `xls_semantic_noop_edit_save`,
  `xls_semantic_one_edit_save`); generated comments-corpus source/eager and
  owned-source open/list/one-cell cases (`xls_source_backed_open`,
  `xls_eager_open_list_worksheets`, `xls_source_backed_open_list_worksheets`,
  `xls_eager_open_one_cell`, `xls_source_backed_open_one_cell`,
  `xls_owned_source_open`, `xls_owned_source_open_list_worksheets`,
  `xls_owned_source_open_one_cell`); named-file simulated-range and owned
  controls for open/list/one-cell/all-cells/full text
  (`xls_range_source_open`, `xls_range_source_open_list_worksheets`,
  `xls_range_source_open_one_cell`, `xls_range_source_open_all_cells`,
  `xls_range_source_open_full_text`, `xls_owned_source_control_open`,
  `xls_owned_source_control_open_list_worksheets`,
  `xls_owned_source_control_open_one_cell`,
  `xls_owned_source_control_open_all_cells`,
  `xls_owned_source_control_open_full_text`); validation
  (`xls_validation_report`); comments one/batch eager/source-backed edits
  (`xls_comments_eager_edit_save`, `xls_comments_source_backed_edit_save`,
  `xls_comments_eager_batch_edit_save`,
  `xls_comments_source_backed_batch_edit_save`); fixed-width Number and
  RK/MulRK eager/source-backed and plan-only edits
  (`xls_numeric_eager_number_edit_save`,
  `xls_numeric_source_backed_number_edit_save`,
  `xls_numeric_eager_rk_mulrk_edit_save`,
  `xls_numeric_source_backed_rk_mulrk_edit_save`,
  `xls_numeric_plan_only_number_edit_save`,
  `xls_numeric_plan_only_rk_mulrk_edit_save`); and one/batch worksheet hide
  eager/source-backed edits (`xls_visibility_eager_edit_save`,
  `xls_visibility_source_backed_edit_save`,
  `xls_visibility_eager_batch_edit_save`,
  `xls_visibility_source_backed_batch_edit_save`).

- **PPT (29):** fresh authoring (`ppt_fresh_write_to`); synthetic open, slide
  list, one shape text, full text, slide-order snapshot, exact no-op and one
  shape-text save (`ppt_semantic_open`, `ppt_semantic_list_slides`,
  `ppt_semantic_one_shape_text`, `ppt_semantic_full_text`,
  `ppt_slide_order_snapshot_open`, `ppt_semantic_noop_edit_save`,
  `ppt_semantic_one_edit_save`); eager/source-backed selected-shape, fresh-open
  selected-shape and repeated selected-shape queries
  (`ppt_source_backed_one_shape_text`, `ppt_semantic_repeated_shape_text`,
  `ppt_source_backed_repeated_shape_text`,
  `ppt_semantic_fresh_open_one_shape_text`,
  `ppt_source_backed_fresh_open_one_shape_text`); named-file simulated-range
  and owned controls (`ppt_range_source_open`,
  `ppt_range_source_open_one_shape_text`, `ppt_owned_source_control_open`,
  `ppt_owned_source_control_open_one_shape_text`); named real-file facade open,
  full text and one slide (`ppt_facade_file_open`, `ppt_facade_file_full_text`,
  `ppt_facade_file_one_slide_text`); one text-edit API path
  (`ppt_text_edit_one_edit_save`); and generated picture open, first-image,
  cached-repeat and all-images eager/source-backed cases
  (`ppt_pictures_eager_open`, `ppt_pictures_source_backed_open`,
  `ppt_pictures_eager_first_image`, `ppt_pictures_source_backed_first_image`,
  `ppt_pictures_eager_cached_repeat`,
  `ppt_pictures_source_backed_cached_repeat`,
  `ppt_pictures_eager_open_all_images`,
  `ppt_pictures_source_backed_open_all_images`).

The shared legacy substrate cases are also callable: `cfb_open`,
`cfb_list_streams`, `cfb_read_one`, borrowed/owned `cfb_create_stream`,
`ole_common_open`, `ole_common_put_stream_publish`,
`ole_common_finish_render`, `ole_common_one_edit_save`, shared open/read and
concurrent reads, selective MiniFAT/FAT legacy/shared and simulated-range
reads, and repeated/bulk/concurrent `cfb_open_stream` cases
(`tools/perf-baseline/src/lib.rs:1245-1286`). They exercise container reads,
stream creation, publication and one generic edit; they are not a complete
stream update/delete CRUD matrix.

The current fixture split matters. `build_writer_corpus` creates every fresh
DOC/XLS/PPT corpus in memory (`tools/perf-baseline/src/lib.rs:15296-15360`):
DOC is paragraphs only (3/512/128), XLS is numeric 1x4x4 or 4x128x16, or 128
one-string sheets for payload-heavy, and PPT is text boxes only (1x2, 12x12,
or 16x8) (`:869-902,15363-15446`). The generated comments and visibility XLS
corpora add opaque CFB streams, but are still generated (`:28184-28264,
29290-29363`). The named range-source/facade families are the current routes
that accept real `.xls`, `.ppt` or `.doc` paths (`:3418-3487,12434-12543`).

## Callable API versus measured coverage

The production API is materially wider than the legacy perf cases:

| Area | Current callable API | Current legacy perf/evidence status | Classification |
| --- | --- | --- | --- |
| DOC body | Replace paragraph/text, paragraph bold and character properties, embedded-object add/remove/display, text/embedded/picture transfer, durable patches, composition and three-way planning; commit and profiled CFB observer hooks (`crates/litchi-doc/src/body_text.rs:930-1318,1639-2180`). | Cases time reads, no-op, one paragraph replacement, and a separate public phase route. No legacy case times formatting, embedded removal/addition, transfer, composition, three-way or durable patch application. | **Harness missing** for the callable operations. The semantic body owner has no paragraph insert/delete/reorder method; `replace_paragraph` is the only paragraph structural edit in this surface (`:1654-1669`). Treat paragraph delete/reorder/merge/split as an API gap for this owner, not as an unqualified missing benchmark. |
| XLS cell values | Set/copy scalar or rich values, clear/remove/insert cells, formulas and styles; insert/delete rows and columns; rename sheets; source-backed commits; patch apply/inverse, transfer and three-way planning (`crates/litchi-xls/src/cell_values/mod.rs:1284-1999,2337-2375,2614-2662,3729-3905`). | Cases time one numeric edit, comments, Number/RK/MulRK, one/batch visibility, no-op, and read selectors. No legacy case times cell clear/remove/insert, formulas/styles, row/column insertion/deletion, sheet rename, copy/transfer, composition or three-way. | **Harness missing**, not unsupported API. |
| PPT slide order | Read hidden/advance/transition/anchor/media state; move/remove slides; hide, advance, transition, shape text(s), anchors, external media path/playback; insert transfer; commit, patches, transfer and three-way (`crates/litchi-ppt/src/slide_order.rs:572-992,1511-2121,2447-2601`). | Cases time slide-order snapshot open and text no-op/one-edit. Real public slide removal exists in retained scratch baselines, but there is no registered legacy move/remove/hide/transition/transfer case. | **Harness missing**, not unsupported API. PPTX boundary cases do not fill this legacy gap. |
| CFB/OLE common | `OleWriter` can create/update/delete streams and create/delete/move storages (`crates/litchi-cfb/src/writer/core.rs:691-714,893-1000,1242-1297`); the common editor can open, update, put streams and finish/commit (`crates/litchi-ole-common/src/object/editor.rs:100-172,318-530,934-977`). | Cases cover open/list/read, stream creation, common put/publish/render/one-edit, shared reads, selective reads and repeated stream reads. No direct low-level update/delete stream/storage or metadata/reorder matrix. | **Harness missing** for callable low-level operations. |

Against the required CRUD categories in `docs/GOAL.md:250-316` and the
checklist (`docs/CRUD_Scenario_Checklist.md:24-81`), the current legacy status is:

- **Read and selective access:** broad and genuinely callable. DOC paragraph
  and text reads, XLS worksheet/cell/full scans and full text, PPT slide/shape/
  picture reads, CFB selective reads, named-file facade routes, and named-file
  range-source controls exist. These are still mostly opt-in and split between
  generated and real-file corpora.
- **Create:** fresh DOC/XLS/PPT writers and exact output checks exist, but the
  three shapes are synthetic authoring paths. They do not establish independent
  producer reopen evidence or opened-document CRUD.
- **Update:** one scalar/text update is measured for each semantic family; XLS
  comments, fixed-width numeric and visibility give additional one/batch
  cases. DOC formatting/embedded/transfer and PPT structural state remain
  unmeasured in the legacy perf harness.
- **Delete, clear, remove, hide:** XLS hide is measured; the production XLS
  cell clear/remove API is not. Retained public PPT real-file slide removal is
  evidence, but not a registered legacy semantic selector. DOC embedded-object
  removal is callable but unmeasured; paragraph deletion is not admitted by
  the body-text owner.
- **Reorder/move, copy, merge/split:** XLS row/column actions and transfer APIs,
  PPT slide move and transfer, and DOC transfer/three-way APIs are callable but
  have no legacy perf cases. DOC paragraph merge/split/reorder is an API gap in
  the audited semantic owner.
- **No-op, patch and conflict:** exact no-op and one-edit patch/inverse checks
  are present for the narrow DOC/XLS/PPT semantic cases. Composition, durable
  patch, foreign/stale conflict and three-way measurements are absent from the
  legacy matrix even where the APIs exist.
- **Validation and preservation:** `xls_validation_report`, CFB checks and
  per-case exact-byte/semantic oracles exist. They do not amount to a complete
  per-format refusal, independent-producer, unknown-content, atomic-failure or
  limit-boundary matrix.

## Representative gaps and stale statements

The biggest representative gap is **opened-document structural CRUD over
producer-shaped legacy files**. The current named real-file routes cover reads;
the registered edit corpora are generated. Retained real-file evidence does
exist: 0617 exercised a real `54016.xls` numeric edit, real
`FloatingPictures.doc` and `NoHeadFoot.doc` paragraph replacements, and real
`45543.ppt` slide removal; 0620 attributed real XLS edits on
`54016.xls`, `WithCustomViews.xls` and `FormulaEvalTestData.xls`. Those are
scratch probes and retained baselines, not a current broad selector matrix.
The current registry therefore lacks representative legacy cases for cell
clear/remove/row or column movement, PPT slide boundaries, DOC embedded or
format edits, and repeated/bulk/failed/patch-composition behavior over those
files.

The fresh-writer evidence is useful but narrower: 0753 reports exact-byte,
synthetic fresh-write ratios of 0.286/0.787/0.883 for DOC, 0.191/0.527/0.910
for PPT, and 0.374/0.915/1.011 for XLS payload-heavy/large/tiny shapes. It
does not price opened-document CRUD, independent producer round trips or
delete/move/copy operations.

The old 0587 gap paragraph in `docs/performance/CRUD_COVERAGE.md:352-375`
should not be read literally today. 0681 superseded its claims about DOC/PPT
facade reads and XLS all-cells/full-text/range selectors
(`CRUD_COVERAGE.md:331-350`); the current enum and 0627/0638 routes confirm
those selectors. That is a **stale inventory statement**, not evidence that
the corresponding current timings are complete. Conversely, the absence of
legacy structural edit selectors and of a broad real-producer edit matrix is
still current. The 0906 registry README explicitly warns that selector counts
do not establish measured CRUD coverage (`tools/perf-baseline/README.md:21-36`).

## One evidence-backed next qualification

Run one current-source **DOC batched validated-render handoff qualification**
over the already admitted real `NoHeadFoot.doc` and `FloatingPictures.doc`
public edit/commit routes. Use the existing DOC commit/CFB observer seams and
retain the normal public lifecycle. This is a qualification gate, not a speedup
claim:

1. Preserve staging validation and final public semantic validation, exact
   no-op bytes, successive edits, failure atomicity, source/policy freshness,
   untouched-stream identity and bounded retained-output/cloning memory.
2. Re-capture the smaller fixture as the stable control and keep the larger
   fixture mandatory; do not infer a benefit from its observer flags.
3. Reject the candidate if the handoff cannot prove the exact source and policy
   binding or if observer/retained-memory cost offsets the measured phase.

This is the strongest next step supported by retained evidence rather than an
unmeasured opportunity: 0728 places common DOC staging at 53.83–57.42% and
finish at 28.95–30.67% of its control lifecycle, while 0729 measures the
actual public finish at 13.00–13.30% on `NoHeadFoot.doc` and 7.68–8.18% on
`FloatingPictures.doc`; 0729 explicitly nominates this bounded pilot and records
the required gates. It does not authorize a generic render cache or an
optimization claim before a fresh qualification.

## Source and evidence hashes

Hashes are SHA-256 of the files read for this audit. The source-context file is
the binding record for the shared head, office-branch identity and source
hashes.

| File | SHA-256 |
| --- | --- |
| `docs/performance/results/change-0906/source-context.json` | `710a84f9a5ec6d4e8693da2c8fd5d6ba55917348df4b7642c8e34ee7b3603b02` |
| `docs/performance/results/change-0906/registry.json` | `f665ef02f335035ed35eaf3f76254a10110868614af699f79dc9e640c649565a` |
| `docs/performance/results/change-0906/registry-replay.json` | `0806372d20bcdf2f106e041d979a8784e9d4aff6a7db0f885eb0ec45d42a778b` |
| `docs/GOAL.md` | `bed4058bb76330daab8ce9d4bceff639ab3fbd7ea06634158bef41b133c4d1f1` |
| `docs/CRUD_Scenario_Checklist.md` | `d3a63f3aa001ad6e3dda7627b6294230ae87b3b5aaa595c45c3327e026dbf02` |
| `tools/perf-baseline/src/lib.rs` | `b62775d903ccb2490c28af61d6c5c95d8ed826afad4f0e7dc3c83bd41b877903` |
| `tools/perf-baseline/src/lib.rs` (current shared checkout) | `f712da842169399d687eb452c7042deed61682a80efdd40b2c3c178af2c036ea` |
| `tools/perf-baseline/README.md` | `7391b23edc75dd87164ef8d1b3f747085d895264d5dbc410bca2b57b2dbab282` |
| `crates/litchi-doc/src/body_text.rs` | `0e8cb6351be3327322882ae4035230cd241aed0b6208cb8589e7d76692d3d786` |
| `crates/litchi-xls/src/cell_values/mod.rs` | `32fd5b9b3dd8141ce081a97609a49383375229034c3504424500c2ce8596b683` |
| `crates/litchi-ppt/src/slide_order.rs` | `b31c5da680a741b4406fa12888eaa7d10d121f4c0c9a241981eeb81ba7580612` |
| `crates/litchi-cfb/src/writer/core.rs` | `a5ead3dded13baa8bdaa350742bfe79185e39cc2f16f86e433701e939fb2b19c` |
| `crates/litchi-ole-common/src/object/editor.rs` | `ee6079567bf85a0ff68a335fb43999d780fab335a09a9cec8b845c16b7b5b319` |
| `docs/performance/0617-cfb-copy-through-writer-design.md` | `6a86a96903b2a37a34b2ce276a755c46669da50df56fc117cbdbbf4bde5bd4b1` |
| `docs/performance/0620-xls-edit-save-attribution.md` | `82745bb0fc08887fb5016f9f8c187df38b06e6a35b1c23fca82b2bade162a323` |
| `docs/performance/0627-ole2-range-source-selectors.md` | `c7be012759e461f4d3f56114d912552e9957220d32a9f035e13cdd3f8a75d77d` |
| `docs/performance/0638-facade-and-ordinary-save-selectors.md` | `3b4a23c7771593b2a7c909231b595e513944f90d47d2f638cb6c37e7256da9ab` |
| `docs/performance/0728-doc-ppt-current-save-baseline.md` | `1a1e855a6f4f34bd8b1537f6ea00648182dc468e90d7276d11489b270709985f` |
| `docs/performance/0729-doc-public-phase-attribution.md` | `fa565fc7f604ecef4c14efab514fa06d1053ff17ba72b14631030ec8d4ca5075` |
| `docs/performance/0748-cfb-overlay-fingerprint-reuse.md` | `288f96850a1c3a29c08fdff8c3e8fe8f53b8129e51d2f92f31264199be9ef54b` |
| `docs/performance/0753-legacy-fresh-writer-text-paths.md` | `de70d02be2c7601c0d414a5ecb08097983c0fd760c2c7534fc6ce830dc5fb219` |
| `docs/performance/CRUD_COVERAGE.md` | `1b9733cb2890f5ccb63997cd7e37e4c2aab3e15218421ec502c33a287d01f0a6` |
