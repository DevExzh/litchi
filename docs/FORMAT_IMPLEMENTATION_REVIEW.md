# Format Implementation Review — `feat/office-format-completeness`

Date: 2026-08-08 (first draft); re-verified and re-graded the same day at HEAD
`1e8af351`.
Scope: every format family in the workspace, reviewed against the source code,
the normative specs under `3rdparty/specs/`, and the ADRs under `docs/adr/`.
Method: two passes. The first pass performed per-format deep audits that read
the actual code, counted `unwrap`/`expect`/`panic!` in production paths,
verified FEATURE_MATRIX claims against the implementation, and ran
`cargo test` / `cargo clippy` where the environment allowed. The second pass
(this revision) independently re-verified every factual claim against the code
at HEAD — re-measuring line, record, test, and panic-macro counts, re-running
`cargo check` / `cargo clippy` / `cargo test` where decisive, and comparing
coverage claims against the spec tables of contents and record enumerations
under `3rdparty/specs/` ([MS-XLS] §2.4, [MS-XLSB] §2.4, [MS-PPT] §2.13.24,
the [MS-DOCX]/[MS-PPTX] extension lists, ODF 1.4 part-3 schema, RTF 1.9.1
appendix B). Documentation claims were treated as unverified until confirmed
in code.

Timing note: commit `db14d6fc` ("fix: close format implementation review
gaps") landed the first draft of this document **together with** fixes for
many of its findings, and later commits (`200d6b97`, `0850a4f4`, `1e8af351`)
added more. Where a first-draft finding was accurate for the pre-fix tree but
has since been fixed, the text says so explicitly instead of repeating the
stale claim.

## Scoring rubric

- **Functional completeness** (0–100): spec coverage across read / write /
  edit, real-file validation, round-trip evidence, test-gate status.
- **API conformance** (0–100): adherence to ADR-0003 (snapshots / edits /
  patches), ADR-0004 (typed semantic API), ADR-0005 (I/O budgets),
  ADR-0006 (panic-free / lossless preservation), and the format-specific
  ADRs 0012–0027.

Two workspace-wide facts depress several scores:

1. The ADR-0003 patch ecosystem — versioned, format-independent,
   deterministic-JSON wire patches, `History<T>`, `ConflictSet`, three-way
   merge — is implemented **nowhere**. Most formats stop at in-memory
   reversible patches. The full in-memory chain (snapshot → transaction →
   commit → source-checked apply → inverse patch) is now complete in XLSX,
   in the flat-level owners of `litchi-odg` (`FlatDrawingPatch`) and
   `litchi-odc` (`FlatChartPatch`), and in the iWork Pages section-text /
   Keynote slide-notes transactions; every other format stops earlier.
2. Rendering, pagination, and field evaluation are explicitly declared
   out of scope in every FEATURE_MATRIX and are not penalized.

## Score summary

| Format | Functional | API conformance | Position |
|---|---|---|---|
| DOCX | 85 | 78 | One of the two most complete crates |
| XLSX | 81 | 85 | Best ADR-0003 conformance; matrix now honest and boundary-locked |
| PPTX | 80 | 75 | Broadest part coverage; writer graph locked on opened files |
| RTF | 80 | 70 | Very wide read; unknown syntax now preserved; build blocked by lint backlog |
| ODT | 78 | 52 | Deepest ODF crate; frontally violates ADR-0003/0006 |
| DOC | 74 | 70 | Near-full read spectrum; main content not editable; limits now exist |
| XLSB | 73 | 78 | Read+write; no cell/style editing on opened workbooks |
| XLS | 72 | 70 | Wide read, create-only writer; clippy gate red |
| PPT | 70 | 64 | Strong read+create; narrow edit surface; clippy badly red |
| ODS | 68 | 62 | Tracked-changes owner is family-best; ADR-0003 deviations persist |
| Keynote | 64 | 71 | Widest iWork editor surface; speaker-notes transactions new |
| Pages | 63 | 71 | Body text now editable via reversible transactions |
| Numbers | 63 | 70 | Wide editor; no package transaction family |
| ODF formula | 68 | 72 | Small but self-consistent |
| ODP | 62 | 58 | Read-strong; real edit API landed; doctests red at HEAD |
| ODC | 48 | 55 | Chart parsed on open; flat axis edits with full in-memory chain |
| Markdown | 45 | 56 | Export-only; core mapping fixed; hyperlinks/images still dropped |
| ODG | 22 | 40 | Flat parser + in-memory patch chain; semantic model unwired |
| ODM | 20 | 40 | Structural validation; no parsing/editing of master documents |
| ODI | 17 | 30 | Flat read parser only |
| ODB | 15 | 30 | Real-file package tests; zero content parsing |
| OTH | 15 | 38 | Substring-level validation only |

---

## Severe findings (cross-format)

These are the issues that most urgently undermine the branch's support claims,
ordered by severity, as verified at HEAD `1e8af351`.

### S1. Build and quality gates: partially red; four first-draft bullets fixed

ADR-0008 requires continuously buildable phases and warning-denied lint gates.
Verified at HEAD:

- **FIXED: the iWork family compiles.** `cargo check -p litchi-iwa-archive
  -p litchi-pages -p litchi-numbers -p litchi-keynote` is clean;
  `crates/litchi-iwa-archive/src/limits.rs:248-269` now constructs
  `soapberry_zip::office::ArchiveLimits` with all 7 fields. The ~2,300 iWork
  tests are runnable.
- **`litchi-rtf` does not compile: 965 rustc-denied lint errors** (916
  `unused_qualifications`, the rest hidden-lifetime / unreachable-pub).
  Because these are rustc lints, not clippy lints, they break plain
  `cargo check`/`cargo build` of the crate and every dependent — the facade
  `rtf`/`office`/`all` features and `litchi-py` (package name of
  `crates/pyo3-litchi`, which is not a workspace member) all fail here.
  The crate's own 1,008 tests cannot run until the backlog is cleared.
  ADR-0008 records this as a pre-existing backlog, still unpaid.
- **`litchi-ppt` clippy: 7,675 errors** on stable 1.97.1 with CI's actual
  command (`cargo clippy --workspace --all-targets --all-features --no-deps --
  -D warnings`, `.github/workflows/rust-ci.yml:211`, no `continue-on-error`);
  5,262 errors lib-only. The count includes **1,817 `unwrap_used` and 224
  `expect_used`** across all targets, both workspace-denied lints with zero
  in-crate allows.
- **Clippy is also red in `litchi-odraw` (160 errors), `litchi-opc` (160),
  and `litchi-cfb` (1)** — without `--no-deps` a workspace clippy run dies on
  these before even reaching `litchi-ppt`, and they also block
  `cargo clippy -p litchi-xls`.
- **FIXED: the three `litchi-xls` integration tests now pass** (10/10 in
  `xls_external_tables.rs` + `xls_row_block_index.rs`; fixed in `db14d6fc`,
  which also added a `CompatibilityProfile` mechanism with defect reporting).
- **FIXED: `tools/check_crate_boundaries.py` is green** ("crate boundaries
  valid for 64 workspace packages and 226 internal dependency declarations,
  14 explicit debt items").
- **FIXED: `crates/litchi/tests/encryption_facade.rs` has its
  `#![cfg(feature = "encryption")]` header** (added in `db14d6fc`).
- **NEW REGRESSION: `cargo test -p litchi-odp` is red at HEAD** — 9 doctests
  in `authoring/mutable.rs` fail to compile (E0433) since `db14d6fc` wired
  the file back into the module tree (see ODP section).

### S2. FEATURE_MATRIX overclaims: all fixed and regression-locked in `db14d6fc`

Every overclaim cited in the first draft was real for the pre-fix tree and is
now corrected in the matrices themselves; the underlying capability gaps are
real but honestly documented:

- `litchi-xlsb` password encryption now reads ❌/❌/❌ (`FEATURE_MATRIX.md:92`)
  with an honest CFB-wrapper note. The crate still has no `encryption` feature
  and no `litchi-crypto` dependency, so the gap is real but no longer
  overclaimed.
- The five `litchi-xlsx` rows were corrected (conditional formatting 🟡/✅/❌,
  hyperlinks 🟡/🟡/❌, defined names 🟡/✅/🟡, page breaks ❌/❌/❌,
  workbook properties 🟡/🟡/🟡 with the `into_plain_opc()` escape documented)
  and are regression-locked by the new `tests/feature_matrix_boundaries.rs`.
  The code truths behind them (no CF package writer, no typed hyperlink
  model, inert `defined_names()` at `workbook/model.rs:470`, untyped
  `rowBreaks`/`colBreaks`) all remain accurate.
- `litchi-xlsb` slicer/timeline now have real
  `Snapshot/Transaction/Commit/Patch` APIs (`slicer/transaction.rs`,
  `timeline/transaction.rs`, `tests/slicer_timeline.rs`), added in `db14d6fc`.
- `litchi-ods` `dde.rs`/`scenario.rs` are now wired into `model/mod.rs:12,18`,
  exposed read-only on the facade, and exercised by `tests/source_features.rs`;
  `rich_text.rs` and `codec/evaluation.rs` were deleted. Caveat:
  `model/mod.rs:2-5` carries a blanket `#![allow(dead_code)]` — several wired
  model modules (conditional_format, sparkline, detective, …) remain
  retained-but-unused pending the parser migration.
- `litchi-markdown/src/lib.rs:12-14` now correctly states that format
  adapters live in the `litchi` umbrella crate's `markdown` module.

### S3. The ODF family split (ADR-0023): regressions largely adjudicated; real leftovers remain

The first draft's counts were off and `db14d6fc` has since resolved most of
the finding. Verified state at HEAD:

- The deferred suites (pre-fix: 13 `.rs` files / 85 tests in
  `litchi-odt`+`litchi-odf-common`, 17 files in `litchi-ods`, all referencing
  deleted pre-split APIs and never compiled) were deleted and partially
  replaced by **active** tests: corpus round-trips, parser hardening, and
  flat-format read/write now run in-build. The replacements are thinner than
  what was deferred (e.g. parser hardening 17 → 4 tests in `litchi-odf-common`,
  flat 22 → 20), so some assurance was genuinely traded away.
- `litchi-odp`'s 1,591-line dead `MutablePresentation` was reworked into a
  765-line `pub(super)` draft type backing the new public
  `authoring::edit::Snapshot` API (16 green tests in
  `tests/presentation_edit.rs`); the stale doc references are gone.
- ODG drawing-style resources were restored
  (`litchi-odf-common::drawing::resources`, ~4,975 lines, with active tests).
- `test-data/odf/corpus/` (15 files) and `test-data/odf/drawing/` (8 files,
  not 7) both have active consumers again.
- Still open: ODB typed schema support was deleted in the split and remains
  absent; query support is reduced to a ~35-line `Query{name, command}` stub.
  ODS conditional formatting / hyperlinks / in-table shapes+images /
  sparklines have no dedicated active feature tests, and `litchi-ods` still
  carries an **orphaned ~6,200-line `src/codec/content/` tree** that is not
  declared in `codec/mod.rs` at all (plus the blanket `allow(dead_code)` in
  `model/mod.rs`).

### S4. ADR-0003 is frontally violated by the ODF text family (stands)

`litchi-odt::Document` exposes **28** `&mut self` methods that mutate the
package directly (`document/package.rs:213-549`: 8 RDF, 9 forms, 11 embedded
object/chart mutations). The remove/move verbs take raw `usize` indices and
return `Result<()>` (add/replace verbs return `Result<usize>`); no
selector-first `Result<Option<_>>` form exists on this root. Scoped
ADR-0003 models do exist — flat `.fodt` paragraph edits
(`flat/mod.rs:357,482,526` with `Edit`/`Commit`/`Patch::apply`/`inverse`),
protection policy, and annotations — but the packaged `Document` root and
`MutableDocument` have no Edit/Commit/Patch model, and `MutableDocument`
(`mutable/model.rs:24`, an attached mutable root retaining
`source_package`) is precisely what ADR-0023 step 5 says must be deleted once
replacements are verified. The same `&mut self` pattern recurs in
`litchi-ods` (12 facade methods, including named-range and RDF edits) and
`litchi-odp`.

### S5. Panic-free discipline (ADR-0006) is unevenly enforced (updated counts)

Production-path `unwrap`/`expect` counts at HEAD, measured with `#[cfg(test)]`
items stripped (first-draft counts in parentheses where they changed):

- `litchi-odt`: **423** (205 unwrap + 218 expect); `litchi-odf-common`:
  **88** (13 + 75) — direct conflict with the workspace `unwrap_used =
  "deny"` / `expect_used = "deny"` lints (`Cargo.toml:67-68`), zero in-crate
  allows. (First draft: 424 / 90 — within measurement noise.)
- `litchi-ppt`: **147** (60 unwrap + 87 expect) + 7 `unreachable!()`, zero
  allows. (First draft: 131.)
- `litchi-xlsx`: **151** (96 unwrap + 55 expect) + 8 `unreachable!()`; the
  unwraps are mostly infallible-by-construction (`write!` into `String`,
  length-checked indexing), so "guarded" is fair.
- `litchi-docx`: **59 unwrap + 34 expect** + 36 `unreachable!()` — the
  largest `unreachable!()` count among measured crates; 30 of the unwraps
  came with the SDT lifecycle commit. (First draft: 28 + 34.)
- `litchi-xlsb`: 13 unwrap + 72 expect + 13 `unreachable!()`; 12 of the 13
  unwraps sit in one file (`host/cells_reader/codec.rs`).
- `litchi-doc`: ~70 (20 unwrap + 50 expect) + 19 `unreachable!()`.
- `litchi-imgconv`: **6 sites** (2 unwrap + 1 expect + 3 `unreachable!()`);
  the two unwraps remain reachable from untrusted EMF path state
  (`emf/svg/buffer.rs:108,118`). (First draft: ~30; fixed in `db14d6fc`.)
- iWork family: exactly **1 guarded unwrap** across all 17 crates — but the
  "panic-free" phrasing omits 50 `.expect()`, 29 `unreachable!()`, and one
  production `panic!` (`litchi-keynote/src/package.rs:311`, feature-gated
  semantic-source branch).
- `litchi-odt` and `litchi-opc` still lack `#![forbid(unsafe_code)]` (neither
  contains actual `unsafe`; hygiene-only gap). `litchi-opc` is otherwise
  clean (0 production panic sites).
- FIXED: `MutableComment::new` no longer stamps `chrono::Utc::now()` —
  comment dates are caller-supplied validated RFC 3339
  (`litchi-docx/src/writer/comment.rs:34-57,109`); a repo-wide grep finds no
  ambient-clock calls in DOCX production code.

### S6. No general open→edit→save path for existing binary Office files (stands; mechanics corrected)

For XLS, PPT, and DOC the writer is create-only and the reader exposes no
save path; editing an existing file is possible only through per-feature
transaction modules — **XLS: 8 (exact); PPT: 14 modules / ~20 editing
surfaces; DOC: ~11 editing surfaces** (first-draft counts of 20 and "13+"
were generous). None of the ~39 surfaces touches ordinary cell values, shape
text, or body paragraph text: there is still no public path to change a cell
value in an existing `.xls`, rewrite a text box in an existing `.ppt`, or
edit a paragraph in an existing `.doc`. This remains the single largest
functional gap against the "full Office CRUD" goal of ADR-0001, and it is
shared by PPTX (opened presentations' writer graph is locked behind
`UnsafeEdit`, `package/model.rs:61-67`, enforced by `tests/edit_guards.rs`).

The first draft's "silent data loss" claim for PPT no longer holds:
`shapes::ShapeContainer::set_text` (`shapes/shape.rs:555`) atomically refuses
mutations on parsed shapes with a typed `MutationError::SourceBound` (shapes
are marked source-bound during parse, with dedicated refusal tests). The
subtler trap is now in XLS: `Worksheet` exposes public mutators
(`add_cell`, `set_dimensions`, `set_sort_info`, …, `worksheet/mod.rs:189-423`)
on a read model that has **no serialization path at all** — mutations there
really are silently droppable.

### S7. Real-file verification: iWork gap closed; crypto/sign/vba/drawingml remain

- **iWork: FIXED.** `test-data/iwork/` now contains Apple-authored native
  fixtures (`pages/basic.pages`, `numbers/basic.numbers`,
  `keynote/basic.key`, plus app-authored `directory/` package oracles) with
  pinned SHA-256s and native reopen verification documented in the README.
  `tests/native_fixture.rs` in all three crates plus
  `litchi-iwa-archive/tests/native_iwa_preservation.rs` open them from path
  and bytes (4/4 pass). The libetonyek QA samples under
  `3rdparty/libreoffice-core` (`Pages_4.pages`, `Keynote_1-6.key`, …) are
  still referenced only in code comments, never in tests.
- `litchi-crypto`: still no real encrypted Office files **wired into the
  crate itself** — note the real encrypted fixtures in test-data
  (`xor-encryption-abc.xls`, `Password_Protected-*.ppt`,
  `password_*_cryptoapi.doc`, …) are exercised by `litchi-xls`, `litchi-ppt`,
  and `litchi-doc`; the gap is specific to `litchi-crypto`'s test wiring.
- `litchi-sign` (keys generated on the fly; no `.pem`/`.p12`/`.pfx` in
  test-data), `litchi-vba` (no real `.docm`/`.xlsm`/`.pptm` fixtures
  anywhere), and `litchi-drawingml` / `litchi-spreadsheet-drawing` /
  `litchi-ograph` (zero test-data references; coverage only indirect via
  host formats) — all verified as stated.

### S8. Governance and documentation gaps: five of seven fixed in `db14d6fc`

- FIXED: `docs/CRUD_Scenario_Checklist.md` now exists and is referenced by
  `docs/adr/README.md:11`. (Nuance: ADR-0008 itself never referenced the
  file — only a "binary CRUD checklist" concept at 0008:1207.)
- FIXED: `litchi-markdown` is now documented in ADR-0024's topology
  (`docs/adr/0024-current-topology.md:492-501`); it remains absent from all
  other ADRs.
- FIXED: ODG/ODC/ODI/ODB/OTH/ODM and the ODF formula crate now have
  FEATURE_MATRIX documents. Still missing: **`litchi-pages`,
  `litchi-numbers`, `litchi-keynote` have none**.
- FIXED: `docs/FEATURE_MATRIX.md:51-54` now correctly says shared ODF lives
  in `litchi-odf-common` / `litchi-odf-formula` and that `litchi-odf` is a
  thin 48-line detector umbrella.
- FIXED: `crates/litchi-xls/README.md` now uses the real `Workbook` export.
- STILL OPEN: ADR-0023 is double-numbered (`0023-odf-family-crate-split` and
  `0023-iwa-index-foundation`); `docs/adr/README.md:38` indexes only the
  former.
- STILL OPEN: `docs/report/odf-iwa-rtf.md` cites pre-split paths that no
  longer exist (`crates/litchi-odf/src/ods/parser.rs`,
  `crates/litchi-rtf/src/error.rs`, `crates/litchi-iwa/build.rs`, …) — partly
  inherent to its genre as a historical split audit.

---

## Per-format details

### DOCX — 85 / 78

Strengths (verified): ~111k lines of Rust (`tokei`: 110,823 code, 133,140
total); every [MS-DOCX] extension family 2.2.1–2.2.13 has a read model except
the umalqura calendar (2.2.7, zero matches — and no `CalendarType` model at
all); within 2.2.3 sdtPr, `appearance`, `color`, `webExtensionsLinked`, and
`webExtensionCreated` are unmodeled while checkbox/repeating-section are read
and authored (`content_control/authoring.rs`, 708 lines); modern comments
(1,176-line codec), conflict revisions (1,858), SDT checksums, OpenType
extensions (949-line codec, 1,701-line module); full package layer (OPC,
Strict/Transitional, MCE, signatures, encryption, read limits); writer covers
paragraphs/tables/sections/revisions/comments/fields/SDT/OLE/VML/SmartArt/
watermarks/TOC. 931 `#[test]` functions, real POI and LibreOffice fixtures, a
fuzz target, and recorded macOS Word open/inspect evidence in ADR-0008.

Deductions (verified): no document-level ADR-0003 model — body editing is
`&mut MutableDocument`, and the 13 part-level `transaction.rs` modules have
no unified patch type; `Block` (`document/model.rs:69-76`) and the sibling
`Element` enum (:60-65) both lack the ADR-0007 `Unknown` fallback and there
is no public `Inline` enum; `litchi-word` (holding the `Visibility`
projection vocabulary) is an orphan crate — declared in workspace deps but
nothing depends on it, so visible/review/all projections are unimplemented;
no `litchi-math` crate (math is raw inert OMML, `src/math.rs`); raw
relationship IDs leak into the public API (`hyperlink.rs:88` `pub fn r_id`,
`package/package/merge.rs:19`) against ADR-0004; 27 unwrap + 34 expect + 36
`unreachable!()` in production (invariant-guarded but nonzero).
~~ambient-time violation~~ FIXED (see S5).

### XLSX — 81 / 85

Strengths (verified): the most complete ADR-0003 implementation in the
workspace — immutable `Arc` snapshot (`workbook/model.rs:136`), `edit()` →
transaction → `Commit{workbook, patch}` → `Patch::inverse()` → source-checked
`Workbook::apply()`, structured `ConflictSet`/`JoinError` (all in
`workbook/edit/model.rs`, re-exported via the 9-line `src/edit.rs`; the first
draft cited the pre-move path); 28 non-raw part-level owners (calc chain per
ADR-0018, data validation, query tables, connections, slicers, timelines,
rich values, a 5,708-line pivot module); ADR-0017 producer templates fully
compliant with byte-parity tests (`xml-minifier/tests/ooxml_assets.rs:232`
over all 8 templates); zero `panic!` in production (8 invariant-guarded
`unreachable!()` remain); 814 tests with exactly 73 real fixtures from three
corpora (poi, ooxml, libreoffice-core). Managed encryption (feature
`encryption`, `tests/encryption.rs`, mutation guards
`ensure_mutation_allowed`/`save_reencrypted*` woven through the edit/apply
paths) is a significant domain the first draft missed.

Deductions (verified): Survey parts ([MS-XLSX] 2.1.9 — `survey` element,
CT_Survey*) entirely absent and undocumented; no dynamic-array spill or
formula evaluation (honestly marked; chartEx exists only as typed-inert
chart-sheet parts, ~4,881 lines); patches have no JSON wire
form/`seal()`/`History` (self-acknowledged debt at
`workbook/edit/model.rs:1401-1406`); `NumberFormat{pub id, pub code}`
accepts anything unvalidated (`style/stylesheet/number_format.rs:8-20`) and
`CellFont.color` is a raw hex `String`, violating ADR-0004/0007; 151 guarded
production unwrap/expect sites. ~~Five FEATURE_MATRIX overclaims~~ FIXED and
boundary-locked (see S2).

### PPTX — 80 / 75

Strengths (verified): broadest part coverage in the workspace — **19 of 20**
[MS-PPTX] extension families 2.2.1–2.2.20 have bounded typed readers (2.2.13
Office App Extensions has none — `webextensionref` appears only in tests; the
2.2.2/2.2.9/2.2.10 extension elements `bmkTgt`/`bounceEnd`,
`modId`/`creationId`, `presenceInfo`/`threadingInfo` are also untyped); ~15
feature domains with high-quality reversible transactions (notes, fonts per
ADR-0022, table styles per ADR-0020, sections, custom shows, guides, designer
tags, comments, OLE/ActiveX, 3D models, zooms, math, tracks, classification);
ADR-0004 exemplars `shape::Scene` and the `time::Offset` decimal-time
grammar; **ADR-0013 is now fulfilled** — `Package::{notes, put_notes,
remove_notes, clear_notes}` use semantic slide selectors backed by a
`notes::{Snapshot, Transaction, Commit, Patch}` source-checked edit model
(`package/model.rs:78-190`, `tests/pptx_notes_facade.rs`); managed package
encryption (`tests/pptx_package_encryption.rs`, `PackageEncryption` policy
checks in `save`/`to_bytes`); 602 tests, exactly 76 real `.pptx` files, ~25
save→reopen test functions.

Deductions (verified): the writer graph of opened presentations is locked —
`presentation_mut()` returns `UnsafeEdit` for opened packages
(`package/model.rs:61-67`, locked by `tests/edit_guards.rs`); slide
add/remove/reorder and shape editing work only on newly created packages
(though ~15 package-level semantic edit domains are fully transactional on
opened packages — the first draft's headline overstated the lock); no
chart/table creation on the facade (chart is read-only inventory, tables
style-only); `litchi-slide` defines the ADR-0007 `LayoutRole`/`Review`/`Look`
vocabulary but nothing depends on it (dead code); the main facade is
`&mut self` + whole-graph clone rollback (`package/codec.rs:262-287`) rather
than the snapshot model outside the notes domain; no pptx-specific fuzz
target.

### XLS (BIFF8) — 72 / 70

Strengths (verified): ~112k lines (112,393 in `src/`); wide typed read
coverage (SST, full styles, pivots, chart metadata, three encryption
profiles, signatures, VBA metadata, revisions, toolbars, XML maps — each
confirmed in `src/`); typed record coverage ~201–209 of the **356** records
in [MS-XLS] §2.4 (**~56–59%**, not "~209/466 ~45%" — 466 matches nothing in
the spec); opaque records not guaranteed to survive typed mutation (honestly
documented, `FEATURE_MATRIX.md:32`); substantial create-only writer with 46
write→reopen test files; ADR-0012 (checked formula references) and ADR-0027
(sheet anchors) fully implemented; exactly 1,165 tests with 37 real fixtures
including encrypted files; a `CompatibilityProfile` mechanism
(`OpenOptions::with_compatibility_profile`, `src/compatibility.rs`) closes
the real-fixture acceptance gap honestly with defect reporting.

Deductions (verified): **clippy is red** (S1: `litchi-odraw` 160 errors +
`litchi-cfb` 1); ~~3 committed tests fail~~ FIXED in `db14d6fc` (10/10 pass);
no general open→edit→save for existing workbooks; 8 per-feature transaction
modules, none touching cell values; write-side formulas are a **27-variant**
Ptg subset (no PtgName/FuncVar/Ref3d/array constants); ADR-0016's migration
debt unpaid — public writer APIs still take raw integers
(`set_column_width(col: u16)`, `freeze_panes(u32)`,
`set_row_height(u32)`); `Worksheet` has public mutators against ADR-0003 —
and with no save path, those mutations are silently droppable (S6); weak XOR
encryption can be written without an explicit policy
(`Writer::set_password(_, EncryptionProfile::XorObfuscation)`), against
ADR-0006 (the "credentials are Clone" sub-claim is weaker than stated: only
the crate-internal `WriterEncryption` derives `Clone`; there is no public
credentials type).

### XLSB — 73 / 78

Strengths (verified): read+write; **539 of 876** record kinds named (~62%;
`raw/kind.rs` has exactly 539 `pub const` kinds; [MS-XLSB] §2.4 has 876
record sections — the first draft's 875 was off by one); **567 tests + 8
doctests green** (11 `vba-inspection`-gated tests not run by default); 11
real fixture files (one also mirrored under `test-data/poi`) including
producer-quirk regression tests; **8 feature domains with source-checked
Snapshot/Transaction/Commit/Patch** (xml_maps, slicer, timeline,
external_link, shared_workbook, comments, connections,
web_extension_bindings) plus sparklines and cell watches via the equivalent
Snapshot/Edit/Commit/Patch pattern (slicer/timeline landed in `db14d6fc`);
consistent laziness for external content; opened workbooks already accept
package-level mutations (`apply_sparklines`, `apply_cell_watches`,
`put_ribbon`, `set_vba`, `set_connections`, sign/resign, `edit_opc`) and
re-save preserves parts byte-wise (`Workbook::save`,
`workbook/package.rs:541`).

Deductions (verified): ~~false encryption matrix claim~~ FIXED (S2);
~~package-level `Error` not `#[non_exhaustive]`~~ FIXED (`host/error.rs:11`,
with a forward-compat regression test); ~~no whole-book preserve re-save~~
INCORRECT — package-level save exists and preserves parts byte-wise (what is
missing is a semantic cell-model round-trip); semantic coverage roughly half
the spec (Data Model, rich values, MDX, smart tags, ActiveX absent — real
spec families: BrtBeginDataModel §2.4.46, BrtBeginRichValueBlock §2.4.195,
BrtBeginMdx* §2.4.106-110, BrtBeginCellSmartTags §2.4.19, BrtActiveX §2.4.4;
mostly declared out of scope); no cell/style transaction editing on opened
workbooks; writer is create-only; 13 production unwraps + 72 expects + 13
`unreachable!()`; the unified facade exposes XLSB read-only.

### PPT (binary) — 70 / 64

Strengths (verified): effectively full RecordType recognition — 208 of the
218 spec values ([MS-PPT] §2.13.24) in the typed `RecordType` enum
(`consts.rs:8`, 210 variants), the remaining 10 handled via typed
raw-constant parsers in dedicated modules; ~63 read accessors on
`Presentation`; a create-from-scratch writer (shapes, rich text, tables,
pictures, animations, sounds; VBA and encryption **feature-gated** behind
`vba-inspection`/`encryption`); ~20 ADR-0003 editing surfaces across 14
transaction modules; **1,097 tests green** on default features (9 ignored)
with 25 real fixtures and byte-exact round-trip assertions; CryptoAPI
encryption read verified against real POI files (write verified by
self-round-trip only, and all encryption tests are feature-gated).

Deductions (verified): no public path for general content edits of existing
files (S6); ~~`Shape` mutators silently lose mutations~~ FIXED — parsed
shapes are source-bound and mutators fail loudly with
`MutationError::SourceBound` (`shapes/shape.rs:555-571`); detached-shape
mutations still have no serialization path; **the "transitions" writer
capability does not exist** — `transition/writer.rs::write_transition` is
orphaned, never called by `Writer`; `Writer::add_chart` unconditionally
returns `UnsupportedAuthoring` (`writer/core/model/semantic.rs:814-849`);
dual-track shapes API (`Box<dyn Shape>` trait objects alongside the
`ShapeEnum` data enum) against ADR-0004; **~390** top-level re-exports (72
`pub use` statements); 147 production unwrap/expect against denied lints and
7,675 clippy errors across all targets (S1).

### DOC (binary) — 74 / 70

Strengths (verified): near-full read spectrum — FIB, piece table, FKP/bin
tables, styles, numbering, all seven stories, ~28 typed field re-exports,
comments, bookmarks, revisions, sections, tables, plus long-tail structures
(mail merge, smart tags with factoid validation, command bars, route slip,
MTEF equations, OLE controls) — all confirmed in `src/`, and no major [MS-DOC]
§2.4–2.9 structure family found wholly absent beyond the listed DOP gaps
(minor unmodelled long-tail: 2007-era OssTheme, DocUndo streams, legacy
`*Old` Bkd/Pgd/Afd variants); a 19,686-line full-stack create writer;
**~11 per-feature editing surfaces** (9 `Editor` structs + story-level
`RevisionEditor` + `property_set::Transaction` — "13+" was an overstatement);
**1,089 tests green**, 40 real fixtures plus POI encrypted files, one fuzz
target, zero todo/unimplemented in `src/`.

Deductions (verified): no unified snapshot/edit/patch model for main content
— the crate is "read-only reader + create-only writer + per-feature
patchers"; patches are binary-stream replacements with no JSON form or
reversibility grading; ~~no memory budgets~~ FIXED in `db14d6fc` — `Limits`
with enforcement now exists (`package/model.rs:12`,
`open_with_limits`/`from_ole_file_with_limits`/`document_with_limits`,
`tests/doc_open_limits.rs`), but defaults are unbounded and `Package::open`
still exposes no limits parameter; no accepted/rejected text projections
(ADR-0007); password passed as plain `Option<&str>`, no `Locked`/`Sensitive`
type-state (ADR-0006); DOP 2007/2010/2013 and DopMth unmodelled; ~70
production unwrap/expect + 19 `unreachable!()`; the FEATURE_MATRIX
implementation map cites pre-refactor paths (`src/package.rs` etc. are now
directories).

### RTF — 80 / 70

Strengths (verified): very wide read coverage — **1,154 typed control-word
variants (~1,490 spellings)** in the `ControlWord` enum/dispatch plus an
`Unknown` fallback (the first draft's "~400" understated this ~3×); a diff
against the RTF 1.9.1 appendix-B index shows ~327 spec-listed words
undispatched (~86% coverage; genuine gaps: smart quotes `\lquote`…,
background-pattern `\bg*`/`\chbg*`, cell-spacing `\clsp*`, color-scheme
mapping, the entire mail-merge `\mm*` family, `\htmltag`/`\htmlrtf`) — these
fall into the `Unknown`/opaque path; nested tables, shapes, legacy drawing,
OLE objects, dozens of typed fields, EQ+OMML math, revisions, compressed RTF
both directions per [MS-OXRTFCP] §2.2/§2.3 (LZFu with the spec-mandated
preloaded dictionary and CRC32 both ways); the immutable snapshot facade is a
textbook ADR-0003/0004 implementation (cheap shared `Arc` handle, lazy
`OnceLock` derived values, `Arc::ptr_eq` snapshot identity, compile-fail
doctests proving immutability); **unknown destinations/control words are now
preserved as bounded opaque nodes and round-tripped byte-for-byte** (added in
`db14d6fc`: `preserve_unknown_destination`/`preserve_unknown_control`,
`src/model/opaque.rs`, writer reinsertion, `tests/opaque_preservation.rs`,
`ParseLimits::with_max_opaque_nodes`), and unmodified documents re-emit the
preserved original bytes via a `preserved_source` fast path; 1,008 tests, 42
named real corpus files, and prefix-truncation / byte-mutation sweeps
(`tests/robustness.rs:139,149`).

Deductions (verified): **the crate does not compile — 965 rustc-denied lint
errors** (S1), so the suite above cannot currently run; the canonical
`RtfWriter::write_document` errors out on structurally-anchored opaque nodes
rather than dropping them (some `skip_group` call sites remain for
known-but-unmodeled group internals); **no Edit/Commit/Patch model at all**,
and `raw::Document` exposes **178** bare `&mut` setters on an attached tree
(against ADR-0001's "mutation is tracked"); no dedicated FEATURE_MATRIX; the
writer `indent` option is dead code; Mac code pages 10001/10007 are rejected
(an intentional, typed, defensible choice).

### ODT (+ common/umbrella) — 78 / 52

Strengths (verified): deepest ODF implementation (111,922 lines in `src/`) —
full package lifecycle, encryption (AES-128/192/256-CBC **and** AES-GCM,
Blowfish-CFB8, Argon2id — exceeding the ODF 1.4 manifest schema's enumerated
PGP/PBKDF2/Blowfish vocabulary via the anyURI escape hatch), signatures
verified against a real LibreOffice XAdES QA file
(`test-data/libreoffice-core/xmlsecurity/.../signed_with_x509certificate_chain.odt`),
RDF, manifest; complete text structures — every claimed item is real ODF 1.4
part-3 vocabulary (`text:ruby`, `text:page-sequence`, `text:tracked-changes`,
`text:section`, `text:table-of-content`, `form:form`, …) mapped to
substantive modules; flat `.fodt` with a scoped Edit/Commit/Patch seam; ~790
`#[test]` functions (plus doctests; "807" is a plausible `cargo test` total);
detection API fully matches ADR-0009 with fuzz targets.

Deductions (verified): see S4 (28 direct `&mut self` package mutations,
`usize`-index edits, `MutableDocument` attached root) and S5 (423 + 88
production unwrap/expect); missing `#![forbid(unsafe_code)]` (hygiene-only);
migration leftovers — two `include!`-based modules (`src/ruby_range.rs:3`,
`src/style/text.rs:158`) and empty git-untracked residue directories
(`src/ods/`, `odf-common/src/migration/` — checkout residue, not committed
content; the first draft's "unreferenced master-document fixture" was wrong:
it is used by `litchi-odm/tests/package_snapshot.rs:10`); ADR-0009's "sole
owner" text not updated after detection moved to `litchi-odf-common`. The
first draft's deferred-test regression (S3) is fixed: those suites now run
in-build, in slimmer form.

### ODS — 68 / 62

Strengths (verified): bounded worksheet graph (repeat-run logical addressing,
merges, coverage, bounded writers); transactional owners for metadata,
calculation settings, annotations, protection, DataPilot, and embedded charts
(but named-range and RDF edits are `&mut self` facade mutations — the first
draft's list was too generous; protection/DataPilot/charts lack Patch types);
the ODF 1.4 tracked-changes owner (exactly 7,335 lines) is the family's best
code — all four record classes (`table:insertion`, `table:deletion`,
`table:movement`, `table:cell-content-change` per part-3 §9.9), limits,
reversible patches, no-op byte preservation, verified against real
LibreOffice files (`change-tracking.ods`, `RecordChangesProtected.ods`);
**flat `.fods` now has a complete `FlatSpreadsheet` owner** (added in
`db14d6fc`: Snapshot/Transaction/Patch/Commit, `src/flat.rs`, 713 lines,
active tests); DDE sources and scenarios are wired and read-only-exposed;
**199 tests green**.

Deductions (verified): conditional formatting (1,173 lines) and sparklines
(856) remain unwired public value types with no facade/worksheet API; the
**~6,200-line `src/codec/content/` tree is orphaned** — not declared in
`codec/mod.rs` at all — and `model/mod.rs` carries a blanket
`#![allow(dead_code)]` masking retained-but-unused modules; `&mut self`
facade edits and the `MutableSpreadsheet` attached root against
ADR-0003/0023; the 17 deferred test files were adjudicated in `db14d6fc`,
but conditional formats / hyperlinks / in-table shapes / sparklines got no
dedicated active replacements.

### ODP — 62 / 58

Strengths (verified): strong read model (23 `DrawingShapeKind` variants
across all `draw:*` shape elements and all six `dr3d:*` elements; animation
`Kind` covers the complete ODF 1.4 part-3 §15 SMIL vocabulary; dedicated
transition/action/settings/page-layout models, 1,175-line layout-master
codec); a from-scratch Builder (1,354 lines); localized edits for
masters/handout/annotations/RDF; **the slide/shape editor is real now** —
the former 1,591-line dead `MutablePresentation` is a 765-line `pub(super)`
implementation behind the public `authoring::edit::Snapshot` API with 16
green tests; real-file tests hard-fail on error (no silent skips remain);
167 unit/integration tests green.

Deductions (verified): **9 doctests in `authoring/mutable.rs` fail to
compile at HEAD (E0433)** — `cargo test -p litchi-odp` is red, a regression
introduced by the fix commit itself; chart transactions require a manual
`commit().into_bytes()` → `from_bytes` republish (deliberate —
"transactions never mutate this presentation" — but still friction); `Shape`
is a pub-field struct with a type tag and `Option<String>` coordinates —
exactly the pattern ADR-0004 forbids; encrypted open exists via shared
`litchi-odf-common` machinery but **no in-crate test exercises an encrypted
ODP**; no forms, encryption write, or signatures; frame-child elements
(`draw:text-box`, `draw:applet`, `draw:plugin`, `draw:floating-frame`) have
no dedicated shape kinds.

### ODG / ODI / ODB, ODC, ODF formula, OTH / ODM

`db14d6fc` grew several of these shells past the first draft's description;
they are no longer isomorphic:

- **ODG — 22 / 40**: still a package shell at the semantic level (model types
  `Layer`/`Page` disconnected from any codec), but now has a real flat-XML
  parser (pages/shapes) and a complete in-memory ADR-0003 chain
  (`FlatDrawingEdit::set_shape_text` → `FlatDrawingCommit` →
  `FlatDrawingPatch` with `inverse()`/`is_applicable_to()`/`apply()`), plus a
  real-content fixture test.
- **ODI — 17 / 30**: flat read parser for frames; no edit surface.
- **ODB — 15 / 30**: still zero content parsing (schema support deleted in
  the split; query reduced to a name/command stub), but now the only shell
  with real binary-file package tests (`tdf132924.odb`, `biblio.odb` from
  LibreOffice's corpus).
- **ODC — 48 / 55**: was a 336-line shell; now 1,011 lines whose facade
  parses the chart tree on open (`chart::read`) and whose flat module
  supports axis edits with the full in-memory patch chain. The real
  capability still lives in `litchi-odf-common::chart` (2,598 lines total:
  1,615 authoring + 582 reader + 401 view/glue — the first draft
  double-counted the reader). Chart class is a free `String` (only
  QName-validated) rather than the part-3 **§19.15** namespaced-token
  vocabulary (12 predefined `chart:*` values; the first draft cited §11).
- **ODF formula — 68 / 72**: small but complete — a **34-kind** validated
  MathML tree (33 MathML kinds + `Other`; "~40" was high) with byte-exact
  round-trip, `.odf`/`.otf` packages, and atomic `set_root` (serialize →
  re-validate → swap, failure leaves the original untouched).
- **OTH — 15 / 38**: still a ~294-line snapshot shell with substring-level
  validation (`xml.contains("<office:text")`), unwired semantic types, and no
  real-file tests.
- **ODM — 20 / 40**: grew to 521 lines — the substring check is now only a
  pre-check before a full NsReader structural pass (root/body/section
  invariants, DOCTYPE rejection, depth caps), with hardening tests and a
  real-content fixture test; still no parsing or editing of master documents.

(`litchi-odraw` is not an ODF crate at all — it is the MS OfficeArt binary
record library, 11,092 lines, and is in good shape on its own terms.)

### iWork (Pages 63 / Keynote 64 / Numbers 63; API 71 / 71 / 70)

Strengths (verified at HEAD `1e8af351`): **the stack compiles** (S1); clean
layered architecture — now 17 crates under ADR-0028 (iwa-monolith-exit),
including new `litchi-iwa-cache`/`litchi-iwa-graph`/`litchi-iwa-index`, with
the legacy structured host seam retired at HEAD; complete physical layer
(ZIP/Snappy/exactly 40 protobuf schemas + 9 hand-written projections); an
extraordinarily wide editor write surface — **PagesEditor: 704 pub fns across
98 impl files including 186 chart editors; KeynoteEditor: 688 pub fns with
full build animations (typed `Effect` model, bounded `Unknown` fallback for
unmodeled Apple effects); NumbersEditor: 653 pub fns with broad
conditional-highlight predicates** (the first draft's 184/~60/224 counts
described single files, and "full predicates with dependency translation" was
wrong on both counts: no duration predicates, and row/column inserts on
volatile-highlight tables are *rejected*, not translated); wire-level
patching that preserves unknown protobuf fields byte-for-byte, with targeted
tests; source-checked package apply (fingerprint + byte + semantic re-check)
and byte passthrough when unmodified; **ADR-0003-shaped reversible
transactions now exist** — Pages section/body text (`SectionTextEdit` with
`SectionTextPatch::inverse()`, `package/section_text.rs`) and Keynote
speaker notes (`SlideNotesEdit`, a 1,464-line module added in `200d6b97`);
**native Apple-authored fixtures** with pinned SHA-256s back integration
tests in all three crates (S7); excellent panic posture (exactly one guarded
unwrap family-wide — though see S5 for the expects/`unreachable!`/one
`panic!` the "panic-free" phrasing omitted).

Deductions (verified): **the wide editor surface lives in legacy `litchi-iwa`
and is unreachable through the semantic crates** — `litchi-pages`/
`litchi-numbers`/`litchi-keynote` do not re-export the editors at all; only
the newer, narrower transaction APIs are; the format-neutral `litchi::iwork`
facade is read-only (but `litchi::pages`/`litchi::keynote` wholesale
re-exports do expose package write transactions at top level — the first
draft's "no write path at the top level" was wrong); Pages body text is
editable only as plain-text splices within single-storage sections (rich
formatting, footnotes, inline objects, section breaks out of scope); atomic
temp+rename save exists only in legacy `litchi-iwa` — the semantic crates
have no save-to-disk; **Numbers has no package transaction family**; no
tracked changes; no encrypted-document support (typed refusal, `.iwpv2`
marker → `ErrorKind::Encrypted`); no formula evaluation (AST only;
`litchi-eval` is not wired); legacy editors remain `&mut self`
clone-staging with no reversible patches; no per-crate FEATURE_MATRIX; the
ADR index's 0023 double-numbering leaves the iwa-index-foundation ADR
unindexed.

### Markdown — 45 / 56

Export-only by design (verified: no Markdown parser anywhere in the
workspace). Most first-draft deductions were fixed in `db14d6fc`: headings
now resolve from DOCX `style_id` and DOC/RTF `outline_level` (`writer.rs:337,
363-387`) and emit `#` prefixes; list detection uses `numbering.xml`
semantics for DOCX/DOC (`elements_with_resolved_list_items`,
`paragraph_list_binding`; unresolvable numbering returns
`Err(Unsupported)`), with the textual heuristic kept only as a fallback for
RTF/ODT/Pages; body text is escaped (`escape.rs`, all output via
`write_literal`); the parallel path propagates errors with `?` instead of
swallowing them; formula conversion failures return `Err` instead of
embedding placeholder strings; the two `unreachable!()` are gone; the
emission layer (now 2,369 lines in `crates/litchi/src/markdown/`) has 29
tests; and ADR-0024:492-501 now documents the crate and blesses the
centralized-adapter topology that `lib.rs:12-14` describes.

Remaining deductions (verified): hyperlinks, images, blockquotes, code
blocks, footnotes, and underline are still unimplemented and silently dropped
(`write_run` handles only bold/italic/strikethrough/script; a stale
`config.rs:30` doc comment still advertises underline); no real-file/golden
validation — all tests build synthetic in-memory packages; ODT/Pages
paragraphs get no heading detection at all even though ODT has outline
levels; dead-code leftovers (`format_formula_placeholder`, `write_list_item`,
`push`/`write_fmt`). The options vocabulary is typed and the leaf crate is
panic-free (zero unwrap/expect/unreachable/panic in `litchi-markdown/src/`).

---

## Overall assessment

First tier — DOCX, XLSX, PPTX, RTF: deep functionality, solid test corpora,
broadly conformant APIs. The common gap remains a unified editing model for
existing documents per ADR-0003 (XLSX excepted); RTF's first-tier scores are
conditional on clearing the 965-error lint backlog that currently blocks its
build. Second tier — XLS/XLSB/DOC/PPT: strong readers and create-only
writers; open→edit→save for main content is missing across the board (S6),
and the clippy gate is red for PPT and XLS's dependency chain. ODF: the
ADR-0023 split's regressions were largely adjudicated in `db14d6fc`, and ODP
gained a real editing API (at the cost of 9 red doctests); ODT is
functionally deep but frontally violates the transactional and panic-free
ADRs. iWork: compiles again, gained native-fixture backing and
ADR-0003-shaped text/notes transactions, but the vast editor surface is still
reachable only through the legacy crate. Markdown and the small ODF families
improved but ODI/ODB/OTH should still not be advertised as supported formats
in their current state.
