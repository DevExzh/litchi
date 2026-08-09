# Format Implementation Review — `feat/office-format-completeness`

Date: 2026-08-08 (first draft); re-verified and re-graded the same day at HEAD
`1e8af351`; re-verified and re-graded again on 2026-08-09 at HEAD `46e88ebc`.
Scope: every format family in the workspace, reviewed against the source code,
the normative specs under `3rdparty/specs/`, and the ADRs under `docs/adr/`.
Method: three passes. The first pass performed per-format deep audits that read
the actual code, counted `unwrap`/`expect`/`panic!` in production paths,
verified FEATURE_MATRIX claims against the implementation, and ran
`cargo test` / `cargo clippy` where the environment allowed. The second pass
independently re-verified every factual claim against the code at `1e8af351`.
The third pass (this revision) re-verified every claim at HEAD `46e88ebc`
after the two defect-closing commits `d1b775d4` and `46e88ebc`, using fourteen
independent per-family audits that re-measured line, record, test, and
panic-macro counts in the live tree, re-ran `cargo check` / `cargo clippy` /
`cargo test` where decisive, and compared coverage claims against the spec
tables of contents and record enumerations under `3rdparty/specs/` ([MS-XLS]
§2.4, [MS-XLSB] §2.4, [MS-PPT] §2.13.24, [MS-PPT] §2.6.6, the
[MS-DOCX]/[MS-PPTX] extension lists, ODF 1.4 part-3 schema §9.9/§15/§19.15,
RTF 1.9.1 appendix B). Documentation claims were treated as unverified until
confirmed in code.

Timing note: commit `db14d6fc` ("fix: close format implementation review
gaps") landed the first draft of this document **together with** fixes for
many of its findings; later commits (`200d6b97`, `0850a4f4`, `1e8af351`) added
more; and `d1b775d4` ("fix: close remaining format review defects") plus
`46e88ebc` ("fix: harden remaining semantic transactions") closed most of what
remained. Where an earlier finding was accurate for the pre-fix tree but has
since been fixed, the text says so explicitly instead of repeating the stale
claim.

## Scoring rubric

- **Functional completeness** (0–100): spec coverage across read / write /
  edit, real-file validation, round-trip evidence, test-gate status.
- **API conformance** (0–100): adherence to ADR-0003 (snapshots / edits /
  patches), ADR-0004 (typed semantic API), ADR-0005 (I/O budgets),
  ADR-0006 (panic-free / lossless preservation), and the format-specific
  ADRs 0012–0029.

Two workspace-wide facts depress several scores:

1. The ADR-0003 patch ecosystem — versioned, format-independent,
   deterministic-JSON wire patches, `History<T>`, `ConflictSet`, three-way
   merge — is implemented **nowhere**. Most formats stop at in-memory
   reversible patches. The full in-memory chain (snapshot → transaction →
   commit → source-checked apply → inverse patch) is now complete in XLSX,
   in the packaged root of `litchi-odt` (`Document::edit()` →
   `transaction::Edit` → `Commit` → `Patch::inverse()`, new in `d1b775d4` /
   `46e88ebc`), in the package- and flat-level owners of `litchi-odg`
   (`FlatDrawingPatch`, `package::Transaction`), in `litchi-odi`
   (`FlatImagePatch`), `litchi-odc` (`FlatChartPatch`), `litchi-odb`
   (schema-catalog transactions), and in the iWork Pages section-text /
   Keynote slide-notes transactions; every other format stops earlier.
2. Rendering, pagination, and field evaluation are explicitly declared
   out of scope in every FEATURE_MATRIX and are not penalized.

## Score summary

| Format | Functional | API conformance | Position |
|---|---|---|---|
| DOCX | 87 | 81 | Most complete crate; sdtPr visual extensions and r_id leaks closed |
| XLSX | 81 | 87 | Best ADR-0003 conformance; ADR-0004 style violations now fixed |
| PPTX | 81 | 76 | Broadest part coverage; task-pane/webextension graph added |
| RTF | 84 | 72 | Very wide read; build blocker cleared, suite green again |
| ODT | 81 | 70 | Deepest ODF crate; packaged root now transactional (S4 resolved) |
| DOC | 75 | 76 | Near-full read spectrum; bounded opens and zeroizing Password landed |
| XLSB | 73 | 78 | Read+write; no cell/style editing on opened workbooks (unchanged) |
| XLS | 73 | 72 | Wide read, create-only writer; mutation traps now fail loudly |
| PPT | 73 | 72 | Strong read+create; transition writer real; clippy gate green |
| ODS | 72 | 66 | Orphan tree deleted; CF/sparkline now inert-but-inventoried |
| ODF formula | 68 | 72 | Small but self-consistent (unchanged) |
| ODP | 65 | 61 | Real edit API; doctests green; own clippy now red |
| Keynote | 64 | 71 | Widest iWork editor surface (unchanged since baseline) |
| Pages | 63 | 71 | Body text editable via reversible transactions (unchanged) |
| Numbers | 63 | 70 | Wide editor; no package transaction family (unchanged) |
| ODC | 60 | 72 | Typed §19.15 chart class; flat axis edits with full chain |
| Markdown | 55 | 64 | Export-only; silent drops converted to typed errors; underline landed |
| ODG | 40 | 60 | Semantic model wired; package-level reversible shape-text edit |
| ODM | 35 | 50 | Real reference/title parsing on real content; no editing |
| ODI | 35 | 50 | Full flat ADR-0003 chain (frame name/source) |
| ODB | 30 | 45 | Typed schema+query catalog parsing tested on real files |
| OTH | 28 | 45 | Structural validation + projection; no-op patch; no real-file tests |

---

## Severe findings (cross-format)

These are the issues that most urgently undermine the branch's support claims,
ordered by severity, as verified at HEAD `46e88ebc`.

### S1. Build and quality gates: every cited gate now green; new residuals surfaced

ADR-0008 requires continuously buildable phases and warning-denied lint gates.
Verified at HEAD:

- **STANDS: the iWork family compiles.** `cargo check -p litchi-iwa-archive
  -p litchi-pages -p litchi-numbers -p litchi-keynote` is clean
  (`Finished dev profile`, zero errors). Correction: `ArchiveLimits`
  (`soapberry-zip/src/office.rs:54-70`) has **6** fields, not 7 as previously
  stated. The ~2,300 iWork tests are runnable.
- **FIXED: `litchi-rtf` compiles — the 965-error backlog is paid.**
  `cargo check -p litchi-rtf --all-targets` is clean (0 `error:` lines), and
  `cargo clippy -p litchi-rtf --all-targets --all-features --no-deps` is also
  clean. `cargo test -p litchi-rtf`: **1,017 passed, 0 failed** across 151
  targets — the suite that previously could not run is green. `d1b775d4`
  touched 231 `litchi-rtf/` files. Stale residue: ADR-0008
  (`docs/adr/0008-migration-and-verification.md:5260`) still describes the
  RTF strict-lint backlog as unpaid.
- **FIXED: `litchi-ppt` clippy — 7,675 → 0.** Under CI's exact scope
  (`cargo clippy -p litchi-ppt --all-targets --all-features --no-deps --
  -D warnings`): zero errors, exit 0. Production `unwrap`/`expect` dropped
  147 → **~18 reasoned sites** (3 unwrap + 15 expect + 1 `panic!`), each under
  an `#[allow(..., reason = "…")]`; ~104 module-level allows sit on
  `#[cfg(test)]` modules. Caveat: with **default features** 3 clippy errors
  remain (`shadow_unrelated` at `src/presentation/package.rs:77`,
  `must_use_candidate` at `src/text_run/model.rs:317,352`) — CI uses
  `--all-features`, so the gate is green, but the default-feature build is
  not lint-clean.
- **FIXED: `litchi-odraw` (160 → 0) and `litchi-opc` (160 → 0) clippy**
  errors; **`litchi-cfb` lib is clean** but its `write_ole` example still has
  exactly 1 error (`print_stderr`/`eprintln!`) under `--all-targets`.
- **STANDS: the three `litchi-xls` integration tests pass** (full suite:
  1,169 passed / 0 failed across 63 binaries).
- **STANDS: `tools/check_crate_boundaries.py` is green** ("crate boundaries
  valid for 64 workspace packages and 226 internal dependency declarations,
  14 explicit debt items" — numbers identical to the baseline).
- **STANDS: `crates/litchi/tests/encryption_facade.rs` has its
  `#![cfg(feature = "encryption")]` header.**
- **FIXED: the `litchi-odp` doctest regression is gone** —
  `cargo test -p litchi-odp --doc`: **21 passed, 0 failed**, including the 9
  `authoring/mutable.rs` doctests that failed E0433 at `1e8af351` (rewritten
  in `d1b775d4` to drive the public `edit::Snapshot` API).
- **NEW RESIDUALS (previously unreported blockers):** with the cited crates
  cleaned, the next red layer surfaced —
  - `litchi-ole-common`: **1,062 clippy errors** (shadowing /
    `must_use_candidate` / etc., workspace-denied pedantic lints) — now the
    crate that kills `cargo clippy -p litchi-xls`;
  - `litchi-odf-common`: **1,084 clippy errors** (including 70
    `expect_used`) — blocks deps-including clippy of `litchi-odp`;
  - `litchi-odp` itself: **46 lib / 47 lib-test clippy errors**
    (`let_underscore_must_use` ×10, wildcard-match ×9, …) in code `d1b775d4`
    touched — it was not red at the baseline.

### S2. FEATURE_MATRIX overclaims: all fixed and regression-locked (stands)

Every overclaim cited in the first draft was real for the pre-fix tree and is
now corrected in the matrices themselves; the underlying capability gaps are
real but honestly documented:

- `litchi-xlsb` password encryption reads ❌/❌/❌ (`FEATURE_MATRIX.md:92`)
  with an honest CFB-wrapper note. The crate still has no `encryption` feature
  and no `litchi-crypto` dependency, so the gap is real but no longer
  overclaimed.
- The five `litchi-xlsx` rows were corrected (conditional formatting 🟡/✅/❌,
  hyperlinks 🟡/🟡/❌, defined names 🟡/✅/🟡, page breaks ❌/❌/❌,
  workbook properties 🟡/🟡/🟡 with the `into_plain_opc()` escape documented)
  and are regression-locked by `tests/feature_matrix_boundaries.rs`. Of the
  code truths behind them, two are now also fixed in code: `NumberFormat` and
  `CellFont.color` are validated typed values (see XLSX section). Still
  accurate: no CF package writer, no typed hyperlink model, inert
  `defined_names()` at `workbook/model.rs:470`, untyped `rowBreaks`/`colBreaks`.
- `litchi-xlsb` slicer/timeline have real `Snapshot/Transaction/Commit/Patch`
  APIs (`slicer/transaction.rs`, `timeline/transaction.rs`,
  `tests/slicer_timeline.rs`).
- `litchi-ods` `dde.rs`/`scenario.rs` are wired into `model/mod.rs:8,14`,
  exposed read-only on the facade (`facade/mod.rs:106,120`), and exercised by
  `tests/source_features.rs`. The blanket `#![allow(dead_code)]` in
  `model/mod.rs` is **gone** (removed in `d1b775d4`; only three targeted
  per-item allows remain) and the orphaned `src/codec/content/` tree was
  deleted (see S3).
- `litchi-markdown/src/lib.rs:12-14` correctly states that format adapters
  live in the `litchi` umbrella crate's `markdown` module; the stale
  `config.rs:30` underline comment was rewritten and is now accurate
  (underline is implemented).

### S3. The ODF family split (ADR-0023): leftover list mostly closed in `d1b775d4`

The first draft's counts were off and the follow-up commits have since
resolved nearly all of the finding. Verified state at HEAD:

- The deferred suites (pre-split references to deleted APIs) were replaced by
  **active** tests: corpus round-trips, parser hardening, and flat-format
  read/write run in-build. The replacements remain thinner than what was
  deferred (parser hardening 17 → 4 tests in `litchi-odf-common`), so some
  assurance was genuinely traded away. One in-repo doc bug: the ODS
  `tests/deferred/README.md` adjudication table claims "`FlatSpreadsheet` no
  longer exists" — it does exist (`src/flat.rs`, 713 lines).
- `litchi-odp`'s dead `MutablePresentation` is a 802-line `pub(super)`
  implementation behind the public `authoring::edit::Snapshot` API (16 green
  tests in `tests/presentation_edit.rs`).
- ODG drawing-style resources are restored and actively tested.
- **FIXED: ODB typed schema support is back** — `model/catalog.rs` (600
  lines) parses a bounded, inert schema + query catalog from `content.xml`
  (`Catalog::parse` with `Limits`; `Table`/`Column`/`TableKind` from
  `db:schema-definition` and `db:table-representations`), `model/query.rs`
  grew 35 → 48 lines (`escape_processing`, validated ctor), and
  `codec/content.rs` is a 189-line NsReader structural validator; facade
  `catalog()`/`catalog_with()` (`facade/mod.rs:74-86`), tested against real
  LibreOffice files (`tests/schema_catalog.rs`, 6 tests, reusing
  `tdf132924.odb`/`biblio.odb`). Forms/reports/connections remain unmodeled.
- **FIXED: the ~6,200-line orphaned `litchi-ods/src/codec/content/` tree is
  deleted** and the blanket `allow(dead_code)` is gone.
- **CHANGED: ODS conditional formatting / hyperlinks / in-table
  shapes+images / sparklines now have dedicated active tests** — but at
  inventory level, not feature CRUD: the new inert `source_features` module
  (722 lines, `Spreadsheet::source_features()` at `facade/mod.rs:253`, with
  its own `SourceFeatureLimits`, DOCTYPE/entity rejection) exposes per-sheet
  `conditional_format_count()` / `sparkline_group_count()` / hyperlink and
  drawing inventories, covered by `tests/source_features.rs` (202 lines) and
  `tests/source_feature_inventory.rs` (48 lines). Typed CF/sparkline
  attach/edit APIs still do not exist.

### S4. ADR-0003 violation by the ODF text family: largely resolved for ODT

**ODT: FIXED (with residuals).** The packaged root now has the ADR-0003
chain the earlier revisions said was missing: `Document::edit()`
(`document/package.rs:191`) → `transaction::Edit` (staged ops,
`MAX_OPERATIONS` bound, `try_reserve` allocation hygiene) → `commit()`
validates the whole batch and publishes one immutable `Snapshot` → `Commit`
(`results()`/`snapshot()`/`patch()`, `transaction.rs:809-856`) → `Patch` with
byte-exact source-checked `apply()` and `inverse()` (`transaction.rs:866-894`).
`src/transaction.rs` is 982 lines covering RDF, forms, embedded
charts/objects/images/resources, scripts, protection, and paragraph
line-break; `tests/packaged_transactions.rs` (291 lines, 6 tests, all green)
proves no-op byte preservation, reversibility on a real LibreOffice-producer
package, and refusal of signed/encrypted envelopes. The **28** `&mut self`
mutators on `Document` still exist (`document/package.rs:268-672`: 8 RDF, 9
forms, 11 embedded object/chart) but are now all
`#[deprecated(note = "use Document::edit()…commit")]` and delegate internally
into the transaction machinery. Residual deviations: remove/move verbs still
take raw `usize` indices (now inside `Edit` too), and pub `MutableDocument`
(`mutable/model.rs:24`) is retained — the ADR-0023 step-5 deletion is still
pending (it is used inside the commit path). Cosmetic oddity: the
deprecations say `since = "0.0.1"` on a 0.0.1 crate.

**ODS/ODP: the pattern persists.** `litchi-ods` still has exactly **12**
`&mut self` facade mutations (`facade/mod.rs`: `update_protection:61`,
`update_tracked_changes:129`, five named-range/definition edits `:321-338`,
six RDF mutations `:417-462`) plus the `MutableSpreadsheet` attached root
(`authoring/mutable.rs:20`), and DataPilot/protection/charts owners have no
`Patch` types (only three `Patch` structs crate-wide: annotations, flat,
tracked-changes). `litchi-odp`'s main-content editing is now transactional,
but its chart transactions still require a manual
`commit().into_bytes()` → `from_bytes` republish.

### S5. Panic-free discipline (ADR-0006): PPT/imgconv debt paid; ODT still the outlier (updated counts)

Production-path `unwrap`/`expect` counts at HEAD, measured with `#[cfg(test)]`
items stripped (earlier counts in parentheses where they changed):

- `litchi-odt`: **432** (214 unwrap + 218 expect; up from 423 — the new
  transaction code added ~9 unwraps); `litchi-odf-common`: **88** (13 + 75) —
  in tension with the workspace `unwrap_used = "deny"` / `expect_used =
  "deny"` lints (`Cargo.toml:67-68`).
- `litchi-ppt`: **FIXED** — down from 147 (60 unwrap + 87 expect) to **~18
  reasoned sites** (3 unwrap + 15 expect + 1 `panic!`), each carrying an
  `#[allow(clippy::expect_used, reason = "…")]`; 7 `unreachable!()` remain.
- `litchi-xlsx`: **151** (96 unwrap + 55 expect) + 8 `unreachable!()`;
  the unwraps are mostly infallible-by-construction (`write!` into `String`,
  length-checked indexing), so "guarded" is fair.
- `litchi-docx`: **27 unwrap + 33 expect + 36 `unreachable!()`** —
  the largest `unreachable!()` count among measured crates. (An earlier
  revision of this section said "59 unwrap"; that figure was wrong — the
  per-format section's ~27 is the accurate measurement.)
- `litchi-xlsb`: 13 unwrap + 72 expect + 13 `unreachable!()`; 12 of the 13
  unwraps sit in one file (`host/cells_reader/codec.rs`).
- `litchi-doc`: 70 (20 unwrap + 50 expect) + 19 `unreachable!()`.
- `litchi-imgconv`: **FIXED, now 4 sites** (1 `expect` at
  `emf/gdi_objects.rs:476` + 3 `unreachable!()` at `codec.rs:659`,
  `emf/svg/converter.rs:1091,1122`, all guarded). The two unwraps reachable
  from untrusted EMF path state (`emf/svg/buffer.rs:108,118`) were replaced
  by total branches in `d1b775d4`, with regression tests
  (`malformed_line_is_not_buffered`,
  `incomplete_pending_state_is_discarded_without_panicking`). Suite: 209
  tests green.
- iWork family: exactly **1 guarded unwrap** across all 17 crates
  (`litchi-iwa/src/keynote/editor/slide_background_wire.rs:62`) — but the
  "panic-free" phrasing omits 50 `.expect()`, 29 `unreachable!()`, and one
  production `panic!` (`litchi-keynote/src/package.rs:311`, feature-gated
  semantic-source branch). All four numbers re-verified exact at HEAD.
- **`#![forbid(unsafe_code)]` gap: FIXED family-wide.** `litchi-odt`
  (`lib.rs:67`), `litchi-odf-common` (:7), `litchi-odf-formula` (:7), and
  `litchi-opc` (:35) all now forbid unsafe code. The six small ODF crates
  (ODG/ODI/ODB/ODM/OTH/ODC) measure **zero** production
  `unwrap`/`expect`/`unreachable!`/`panic!` sites.
- STANDS: `MutableComment::new` no longer stamps `chrono::Utc::now()` —
  comment dates are caller-supplied validated RFC 3339
  (`litchi-docx/src/writer/comment.rs:34-57,109`); no ambient-clock calls in
  DOCX production code.

### S6. No general open→edit→save path for existing binary Office files (stands; XLS trap fixed)

For XLS, PPT, and DOC the writer is create-only and the reader exposes no
save path; editing an existing file is possible only through per-feature
transaction modules — **XLS: 8 (exact); PPT: 14 modules / exactly 20 editing
surfaces; DOC: exactly 11 editing surfaces** (9 `Editor` structs +
`RevisionEditor` + `property_set::Transaction`). None of the ~39 surfaces
touches ordinary cell values, shape text, or body paragraph text: there is
still no public path to change a cell value in an existing `.xls`, rewrite a
text box in an existing `.ppt`, or edit a paragraph in an existing `.doc`.
This remains the single largest functional gap against the "full Office CRUD"
goal of ADR-0001, and it is shared by PPTX (opened presentations' writer
graph is locked behind `UnsafeEdit`, `package/model.rs:148-154`, enforced by
`tests/edit_guards.rs:11-27`).

Both "silent data loss" traps identified earlier are now closed:

- PPT: `shapes::ShapeContainer::set_text` atomically refuses mutations on
  parsed shapes with a typed `MutationError::SourceBound` (`ensure_mutable`
  at `shapes/shape.rs:622-636`, refusal tests at :850-908).
- **XLS: FIXED in `d1b775d4`.** `Worksheet` now carries `source_bound: bool`
  (`worksheet/mod.rs:36`, set by `mark_source_bound()` on decode,
  `workbook/codec/semantic/worksheet.rs:1001`), and all public `&mut self`
  mutators return `Result` and refuse decoded worksheets with
  `Error::SourceBoundWorksheetMutation` (`add_cell` :198, `set_dimensions`
  :207, `add_merged_cells` :299, `set_autofilter_info` :335,
  `add_autofilter_column` :345, `set_sort_info` :361, `add_pivot_table` :461,
  `protection_mut` :541), regression-locked by `tests/xls_mutation_policy.rs`
  (3 tests). Mutations are no longer silently droppable — they fail loudly.
  (PPT detached-shape mutations still have no serialization path.)

### S7. Real-file verification: iWork gap closed; crypto/sign/vba/drawingml remain

- **iWork: STANDS (fixed earlier).** `test-data/iwork/` contains
  Apple-authored native fixtures with pinned SHA-256s and native reopen
  verification; `tests/native_fixture.rs` in all three crates plus
  `litchi-iwa-archive/tests/native_iwa_preservation.rs` pass (5/5 at HEAD).
  The libetonyek QA samples under `3rdparty/libreoffice-core` are still
  referenced only in code comments, never in tests.
- `litchi-crypto`: still no real encrypted Office files **wired into the
  crate itself** — the real encrypted fixtures in test-data
  (`xor-encryption-abc.xls`, `Password_Protected-*.ppt`,
  `password_*_cryptoapi.doc`, …) are exercised by `litchi-xls`, `litchi-ppt`,
  and `litchi-doc`; the gap is specific to `litchi-crypto`'s test wiring.
- `litchi-sign` (keys generated on the fly; no `.pem`/`.p12`/`.pfx` in
  test-data), `litchi-vba` (no real `.docm`/`.xlsm`/`.pptm` fixtures
  anywhere), and `litchi-drawingml` / `litchi-spreadsheet-drawing` /
  `litchi-ograph` (zero test-data references; coverage only indirect via
  host formats) — all verified as stated.
- Small improvements elsewhere: Markdown gained its first real-fixture test
  (`testPictures.doc` inline-image refusal), ODM gained a real LibreOffice
  master-document fixture test, and ODB's schema catalog is tested against
  real `.odb` files — but ODI/OTH still have zero real-file tests.

### S8. Governance and documentation gaps: ADR-0023 numbering fixed; iWork matrices still missing

- STANDS (fixed): `docs/CRUD_Scenario_Checklist.md` exists and is referenced
  by `docs/adr/README.md:11`.
- STANDS (fixed): `litchi-markdown` is documented in ADR-0024's topology
  (`docs/adr/0024-current-topology.md:492-501`); it remains absent from all
  other ADRs.
- STANDS (fixed): ODG/ODC/ODI/ODB/OTH/ODM and the ODF formula crate have
  FEATURE_MATRIX documents. Still missing: **`litchi-pages`,
  `litchi-numbers`, `litchi-keynote` have none** (re-verified at HEAD).
- STANDS (fixed): `docs/FEATURE_MATRIX.md:51-54` correctly describes shared
  ODF in `litchi-odf-common` / `litchi-odf-formula` and `litchi-odf` as a
  thin 48-line detector umbrella (`crates/litchi-odf/src/lib.rs` is exactly
  48 lines, re-exports only).
- STANDS (fixed): `crates/litchi-xls/README.md` uses the real `Workbook`
  export; `litchi-doc`'s README/FEATURE_MATRIX implementation map was
  rewritten in `d1b775d4` and every cited path now exists.
- **FIXED: ADR-0023 is no longer double-numbered** — the IWA index record
  was renamed to `docs/adr/0029-iwa-index-foundation.md` and is indexed at
  `docs/adr/README.md:44` with the correction noted at :50-51.
- **CHANGED (mitigated):** `docs/report/odf-iwa-rtf.md` still cites 12
  pre-split paths that no longer exist, but `d1b775d4` added a prominent
  "Historical audit record" disclaimer at the top, converting the defect into
  an explicit genre statement.
- STILL OPEN: ADR-0008 (`docs/adr/0008-migration-and-verification.md:5260`)
  still describes the RTF strict-lint backlog as unpaid (paid in `d1b775d4`),
  and ADR-0009's "sole owner" text (`docs/adr/0009:22`) still names
  `litchi-odf::detect` though detection lives in `litchi-odf-common::detect`.

---

## Per-format details

### DOCX — 87 / 81

Strengths (verified): ~115k lines of Rust (`tokei`: 114,853 code, 133,716
total — grew ~4k with the two fix commits); every [MS-DOCX] extension family
2.2.1–2.2.13 has a read model except the umalqura calendar (2.2.7, zero
matches — and no `CalendarType` model at all); within 2.2.3 sdtPr, the
`appearance`/`color`/`webExtensionsLinked`/`webExtensionCreated` gap is
**closed** — typed `Appearance`, `SdtColor`, and inert `WebExtensionBinding`
(`content_control.rs:85-175`) are read, stored on `ContentControl`, authored
via `AuthoringView::{appearance,color,web_extension_linked,
web_extension_created}`, and regression-locked by
`tests/sdt_visual_extensions.rs` (3 tests incl. invalid-value rejection);
checkbox/repeating-section are read and authored (`content_control/
authoring.rs`, 780 lines); modern comments (1,176-line codec), conflict
revisions (1,858), SDT checksums, OpenType extensions (949-line codec,
1,701-line module); full package layer (OPC, Strict/Transitional, MCE,
signatures, encryption, read limits); writer covers paragraphs/tables/
sections/revisions/comments/fields/SDT/OLE/VML/SmartArt/watermarks/TOC. 938
`#[test]` functions, real POI and LibreOffice fixtures, a fuzz target, and
recorded macOS Word open/inspect evidence in ADR-0008. New in `46e88ebc`:
SDT transaction hardening (`content_control/package.rs`, now 2,147 lines —
mutation-index reconciliation, bounded `unsigned_signature_token`,
`try_reserve_edits` preflight) and custom-XML data-store coverage
(`package/package/data_stores.rs`, `tests/custom_xml_data_store.rs`).

Deductions (verified): no document-level ADR-0003 model — body editing is
`&mut MutableDocument`, and the 13 part-level `transaction.rs` modules have
no unified patch type; `Block` (`document/model.rs:69-76`) and the sibling
`Element` enum (:60-65) both lack the ADR-0007 `Unknown` fallback and there
is no public `Inline` enum; `litchi-word` (holding the `Visibility`
projection vocabulary) is an orphan crate — declared in workspace deps but
nothing depends on it, so visible/review/all projections are unimplemented;
no `litchi-math` crate (math is raw inert OMML, `src/math.rs`, 605 lines);
~~raw relationship IDs leak into the public API~~ FIXED in `d1b775d4` —
`Hyperlink::r_id` deleted, `Package::mail_merge_target` takes a validated
opaque `RelationshipId` newtype (`mail_merge/model.rs:23-52`); 27 unwrap +
33 expect + 36 `unreachable!()` in production (invariant-guarded but
nonzero).

### XLSX — 81 / 87

Strengths (verified): the most complete ADR-0003 implementation in the
workspace — immutable `Arc` snapshot (`workbook/model.rs:137`), `edit()` →
transaction → `Commit{workbook, patch}` → `Patch::inverse()` → source-checked
`Workbook::apply()`, structured `ConflictSet`/`JoinError` (all in
`workbook/edit/model.rs`, re-exported via the 9-line `src/edit.rs`); 28
non-raw part-level owners (calc chain per ADR-0018, data validation, query
tables, connections, slicers, timelines, rich values, a 5,708-line pivot
module); ADR-0017 producer templates fully compliant with byte-parity tests
(`xml-minifier/tests/ooxml_assets.rs:232` over all 8 templates); zero
`panic!` in production (8 invariant-guarded `unreachable!()` remain); **816
tests** with exactly **75** real fixtures from three corpora (poi, ooxml,
libreoffice-core). Managed encryption (feature `encryption`,
`tests/encryption.rs`, mutation guards `ensure_mutation_allowed`/
`save_reencrypted*` woven through the edit/apply paths). **The two ADR-0004
style violations are fixed in `d1b775d4`:** `NumberFormat` fields are now
private with validating `new()`/`custom()` (255-char Excel limit, reserved-ID
range, `InvalidNumberFormat` error) and lossless `from_raw()`; `CellFont.
color` is now `Option<FontColor>` (`style/stylesheet/font.rs:333`) —
`FontColorKind{Rgb, Theme, Indexed, Auto, Default}` with checked `Tint`
makes conflicting color bases unrepresentable.

Deductions (verified): Survey parts ([MS-XLSX] 2.1.9 — `survey` element,
CT_Survey*) entirely absent and undocumented; no dynamic-array spill or
formula evaluation (honestly marked; chartEx exists only as typed-inert
chart-sheet parts, ~4,928 non-test lines); patches have no JSON wire
form/`seal()`/`History` (self-acknowledged debt at
`workbook/edit/model.rs:1401-1406`); 151 guarded production unwrap/expect
sites.

### PPTX — 81 / 76

Strengths (verified): broadest part coverage in the workspace — **19 of 20**
[MS-PPTX] extension families 2.2.1–2.2.20 have bounded typed readers, and
2.2.13 is now **partially closed**: the new task-pane facade
(`Package::{task_panes, task_panes_patch, clear_task_panes_patch,
apply_task_panes_patch}`, `package/model.rs:74-128`, added in `d1b775d4`)
re-exports a typed MS-OWEXML model (`AddIn/Panes/Patch/Snapshot/Store`,
`src/web.rs`) with source-checked reversible patches and 3 facade tests
(`tests/pptx_task_panes.rs`) — though the literal slide-tree content-app
`webextensionref` element remains untyped; ~16 feature domains with
high-quality reversible transactions (notes, fonts per ADR-0022, table
styles per ADR-0020, sections, custom shows, guides, designer tags, comments,
OLE/ActiveX, 3D models, zooms, math, tracks, classification, task panes);
ADR-0004 exemplars `shape::Scene` and the `time::Offset` decimal-time
grammar; the notes facade uses semantic slide selectors backed by a
`notes::{Snapshot, Transaction, Commit, Patch}` source-checked edit model
(`package/model.rs:165-283`, `tests/pptx_notes_facade.rs`); managed package
encryption (`tests/pptx_package_encryption.rs`, `PackageEncryption` policy
checks in `save`/`to_bytes`); **604 tests**, exactly 76 real `.pptx` files,
~25 save→reopen test functions.

Deductions (verified): the writer graph of opened presentations is locked —
`presentation_mut()` returns `UnsafeEdit` for opened packages
(`package/model.rs:148-154`, locked by `tests/edit_guards.rs`); slide
add/remove/reorder and shape editing work only on newly created packages
(though ~16 package-level semantic edit domains are fully transactional on
opened packages); the 2.2.2/2.2.9/2.2.10 extension elements
`bmkTgt`/`bounceEnd`, `modId`/`creationId`, `presenceInfo`/`threadingInfo`
remain untyped; no chart/table creation on the facade (chart is read-only
inventory, tables style-only); `litchi-slide` defines the ADR-0007
`LayoutRole`/`Review`/`Look` vocabulary but nothing depends on it (dead
code); the main facade is `&mut self` + whole-graph clone rollback
(`package/codec.rs:310-335`) rather than the snapshot model outside the
notes domain; no pptx-specific fuzz target.

### XLS (BIFF8) — 73 / 72

Strengths (verified): ~112.5k lines (112,495 in `src/`); wide typed read
coverage (SST, full styles, pivots, chart metadata, three encryption
profiles, signatures, VBA metadata, revisions, toolbars, XML maps); typed
record coverage ~201–209 of the **356** records in [MS-XLS] §2.4 (~56–59%);
opaque records not guaranteed to survive typed mutation (honestly documented,
`FEATURE_MATRIX.md:32`); substantial create-only writer with 46 write→reopen
test files; ADR-0012 (checked formula references) and ADR-0027 (sheet
anchors) fully implemented; **1,168 tests** (1,169 pass across 63 binaries,
0 failed) with 37 real fixtures including encrypted files; a
`CompatibilityProfile` mechanism (`OpenOptions::with_compatibility_profile`,
`src/compatibility.rs`) closes the real-fixture acceptance gap honestly with
defect reporting. **New in `d1b775d4`:** source-bound worksheet mutation
guards (S6) and a weak-encryption policy gate — `Writer::set_password(_,
EncryptionProfile::XorObfuscation)` now returns
`Err(Error::WeakEncryptionRequiresExplicitPolicy)`
(`writer/core/codec/lifecycle.rs:76-78`); XOR write requires the explicit
`#[must_use]` `WeakEncryptionPolicy::allow_xor_obfuscation()` token via
`set_xor_obfuscation_password` (:88-94, re-exported `lib.rs:364`).

Deductions (verified): **clippy is still red** but the blocker moved — with
`litchi-odraw` clean, `cargo clippy -p litchi-xls` now dies on
**`litchi-ole-common` (1,062 errors)**, a previously unreported dependency
blocker; the `litchi-cfb` example lint (1 error) is also unfixed; no general
open→edit→save for existing workbooks; 8 per-feature transaction modules,
none touching cell values; write-side formulas are a **26-variant** Ptg
subset (`writer/formula.rs:151`; no PtgName/FuncVar/Ref3d/array constants —
an earlier revision said 27, off by one); ADR-0016's migration debt unpaid —
public writer APIs still take raw integers (`set_column_width(col: u16)`,
`freeze_panes(u32)`, `set_row_height(u32)`); ~~`Worksheet` public mutators
silently droppable~~ FIXED (S6); ~~weak XOR encryption writable without an
explicit policy~~ FIXED (above). The "credentials are Clone" sub-claim
remains weaker than stated: only the crate-internal `WriterEncryption`
derives `Clone`; there is no public credentials type.

### XLSB — 73 / 78 (unchanged; crate byte-identical to the baseline)

Strengths (verified): read+write; **539 of 876** record kinds named (~62%;
`raw/kind.rs` has exactly 539 `pub const` kinds; [MS-XLSB] §2.4 has exactly
876 record sections); **567 tests + 8 doctests green** (11
`vba-inspection`-gated tests not run by default; the feature-gated path was
also run: 586 tests pass); 11 real fixture files (one also mirrored under
`test-data/poi`) including producer-quirk regression tests; **8 feature
domains with source-checked Snapshot/Transaction/Commit/Patch** (xml_maps,
slicer, timeline, external_link, shared_workbook, comments, connections,
web_extension_bindings) plus sparklines and cell watches via the equivalent
Snapshot/Edit/Commit/Patch pattern; consistent laziness for external content;
opened workbooks already accept package-level mutations (`apply_sparklines`,
`apply_cell_watches`, `put_ribbon`, `set_vba`, `set_connections`,
sign/resign, `edit_opc`) and re-save preserves parts byte-wise
(`Workbook::save`, `workbook/package.rs:541`; untouched parts retain source
bytes, `litchi-opc/src/package.rs:92`).

Deductions (verified): semantic coverage roughly half the spec (Data Model,
rich values, MDX, smart tags, ActiveX absent — real spec families:
BrtBeginDataModel §2.4.46, BrtBeginRichValueBlock §2.4.195, BrtBeginMdx*
§2.4.106-110, BrtBeginCellSmartTags §2.4.19, BrtActiveX §2.4.4; mostly
declared out of scope); no cell/style transaction editing on opened
workbooks; writer is create-only; 13 production unwraps + 72 expects + 13
`unreachable!()`; the unified facade exposes XLSB read-only. Neither fix
commit touched this crate.

### PPT (binary) — 73 / 72

Strengths (verified): effectively full RecordType recognition — 208 of the
218 spec values ([MS-PPT] §2.13.24) in the typed `RecordType` enum
(`consts.rs:8`, 210 variants; verified by value against the spec table — the
remaining 10 are handled via typed raw-constant parsers in dedicated
modules, e.g. `COLOR_SCHEME_ATOM_TYPE` at `color_scheme.rs:11`); 62 read
accessors on `Presentation`; a create-from-scratch writer (shapes, rich
text, tables, pictures, animations, sounds; VBA and encryption
**feature-gated** behind `vba-inspection`/`encryption`); exactly 20 ADR-0003
editing surfaces across 14 transaction modules; **1,104 tests green** on
default features (9 ignored) with 28 real fixtures and byte-exact round-trip
assertions; CryptoAPI encryption read verified against real POI files (write
verified by self-round-trip only, all encryption tests feature-gated).
**The transition writer is real now** (fixed in `d1b775d4`):
`transition/writer.rs::write_transition` is called from
`src/writer/core/package.rs:210` via `build_slide_info_atom`, exposed as
`Writer::set_slide_transition`, with authored→read round-trip tests
(`tests/ppt_transition_writer.rs:49-88`); direction bytes were corrected
from LibreOffice-derived values to [MS-PPT] §2.6.6-exact inverses of the
reader (`transition/writer.rs:112-147`).

Deductions (verified): no public path for general content edits of existing
files (S6); parsed shapes are source-bound and mutators fail loudly
(`ensure_mutable`, `shapes/shape.rs:622-636`), but detached-shape mutations
still have no serialization path; `Writer::add_chart` still unconditionally
returns `UnsupportedAuthoring` (`writer/core/model/semantic.rs:947-989`,
hardened with validation but refusing at :986); dual-track shapes API
(`Box<dyn Shape>` trait objects alongside the `ShapeEnum` data enum) against
ADR-0004; ~388 top-level re-exports (72 `pub use` statements); ~~147
production unwrap/expect + 7,675 clippy errors~~ FIXED (S1/S5) — now ~18
reasoned panic sites and a green CI-scope clippy gate, modulo 3
default-feature-only clippy errors.

### DOC (binary) — 75 / 76

Strengths (verified): near-full read spectrum — FIB, piece table, FKP/bin
tables, styles, numbering, all seven stories, **31** typed field re-exports,
comments, bookmarks, revisions, sections, tables, plus long-tail structures
(mail merge, smart tags with factoid validation, command bars, route slip,
MTEF equations, OLE controls); a 19,686-line full-stack create writer;
exactly **11 per-feature editing surfaces** (9 `Editor` structs + story-level
`RevisionEditor` + `property_set::Transaction`); **1,100 tests green**, 40
real fixtures plus POI encrypted files, one fuzz target, zero
todo/unimplemented in `src/`. **ADR-0005/0006 hardening landed in
`d1b775d4`:** `Limits::default()` is now finite (128 MiB package / 64 MiB
stream / 96 MiB aggregate, `package/model.rs:32-37`, with hard ceilings
retained), new `PackageOpenOptions` + `Package::open_with(path, options)`
(`package/codec.rs:36`); passwords are now an owned, non-`Clone`,
`Zeroizing<String>`-backed `Password` type with redacted `Debug`
(`package/model.rs`, exported `lib.rs:111`), replacing the plain
`Option<&str>`; `document_with_options_and_limits` combines supplied limits
component-wise with package limits, retaining the stricter per dimension
(`codec.rs:154-160`); README/FEATURE_MATRIX document all of it.

Deductions (verified): no unified snapshot/edit/patch model for main content
— the crate is "read-only reader + create-only writer + per-feature
patchers"; patches are binary-stream replacements with no JSON form or
reversibility grading (`property_set::Patch` has no JSON; the
tracked-revision commit pairs with a reversible patch per-feature); no
accepted/rejected text projections (ADR-0007) — now explicitly documented as
a deliberate choice in README/FEATURE_MATRIX; DOP 2007/2010/2013 and DopMth
unmodelled (spec-confirmed: [MS-DOC] §2.7.17); ~70 production unwrap/expect
+ 19 `unreachable!()`; the `Password` type is not a full type-state machine,
but the credential-hygiene substance of the deduction is closed.

### RTF — 84 / 72

Strengths (verified): very wide read coverage — **1,156 typed control-word
variants (~1,490–1,520 spellings)** in the `ControlWord` enum/dispatch plus
an `Unknown` fallback; a diff against the RTF 1.9.1 appendix-B index shows
~86% coverage (genuine gaps: smart quotes `\lquote`…, background-pattern
`\bg*`/`\chbg*`, cell-spacing `\clsp*`, color-scheme mapping,
`\htmltag`/`\htmlrtf`) — these fall into the `Unknown`/opaque path.
Correction: the earlier claim that "the entire mail-merge `\mm*` family" is a
gap was **wrong** — 30 `\mm*` spellings are dispatched (`mmconnectstr`,
`mmdatasource`, the full `mmodso*` set) with a typed model in
`src/metadata/mail_merge.rs`; nested tables, shapes, legacy drawing, OLE
objects, dozens of typed fields, EQ+OMML math, revisions, compressed RTF
both directions per [MS-OXRTFCP] §2.2/§2.3 (LZFu with the spec-mandated
preloaded dictionary and CRC32 both ways); the immutable snapshot facade is a
textbook ADR-0003/0004 implementation (cheap shared `Arc` handle, lazy
`OnceLock` derived values, `Arc::ptr_eq` snapshot identity, compile-fail
doctests proving immutability); unknown destinations/control words are
preserved as bounded opaque nodes and round-tripped byte-for-byte
(`src/model/opaque.rs`, writer reinsertion, `tests/opaque_preservation.rs`,
`ParseLimits::with_max_opaque_nodes`), and unmodified documents re-emit the
preserved original bytes via a `preserved_source` fast path; 1,008 `#[test]`
functions (**1,017 pass** including doctests, verified running at HEAD), 44
named real corpus files, and prefix-truncation / byte-mutation sweeps
(`tests/robustness.rs:139,149`). `d1b775d4` additionally expanded typed
models (fields +692 lines, tables +422, document model +439).

Deductions (verified): ~~the crate does not compile~~ FIXED — the 965-error
lint backlog is paid; `cargo check`/`clippy`/`test` are all clean and green
(S1); the canonical `RtfWriter::write_document` errors out on
structurally-anchored opaque nodes rather than dropping them (11 `skip_group`
call sites remain for known-but-unmodeled group internals); **no
Edit/Commit/Patch model at all**, and `raw::Document` exposes **174** bare
`&mut` setters on an attached tree (against ADR-0001's "mutation is
tracked"); no dedicated FEATURE_MATRIX; the writer `indent` option is dead
code; Mac code pages 10001/10007 are rejected (an intentional, typed,
defensible choice).

### ODT (+ common/umbrella) — 81 / 70

Strengths (verified): deepest ODF implementation (113,150 lines in `src/`) —
full package lifecycle, encryption (AES-128/192/256-CBC **and** AES-GCM,
Blowfish-CFB8, Argon2id — exceeding the ODF 1.4 manifest schema's enumerated
PGP/PBKDF2/Blowfish vocabulary via the anyURI escape hatch), signatures
verified against a real LibreOffice XAdES QA file, RDF, manifest; complete
text structures — every claimed item is real ODF 1.4 part-3 vocabulary
(`text:ruby`, `text:page-sequence`, `text:tracked-changes`, `text:section`,
`text:table-of-content`, `form:form`, …) mapped to substantive modules; flat
`.fodt` with a scoped Edit/Commit/Patch seam; **the packaged root now has a
complete, tested ADR-0003 transaction chain** (S4 — `Document::edit()` →
`transaction::Edit` → `Commit` → source-checked reversible `Patch`, 982-line
`src/transaction.rs`, 6 green integration tests); **the chart class is now a
typed ODF 1.4 §19.15 vocabulary** — `litchi-odf-common/src/chart/class.rs`
(235 lines): `ChartClass`/`ChartClassKind` with all 12 predefined `chart:*`
values plus namespaced `Extension` and retained `Unknown`, used by reader and
authoring model; 796 `#[test]` functions; detection API fully matches
ADR-0009 with fuzz targets; `#![forbid(unsafe_code)]` now present
(`lib.rs:67`, also `odf-common:7`, `odf-formula:7`).

Deductions (verified): 432 + 88 production unwrap/expect sites (S5) — the
crate family remains the panic-discipline outlier; residual ADR-0003
deviations (S4): `usize`-index remove/move verbs, retained `MutableDocument`
attached root pending ADR-0023 step 5, deprecated-but-present 28 legacy
mutators; migration leftovers — two `include!`-based modules
(`src/ruby_range.rs:3`, `src/style/text.rs:158`) and empty git-untracked
residue directories (`src/ods/`, `odf-common/src/migration/`, plus new empty
`odf-common/tests/deferred/{cross_family,flat}`); ADR-0009's "sole owner"
text not updated after detection moved to `litchi-odf-common` (S8); the
`litchi-odf` umbrella is a thin 48-line detector re-export (verified exact).

### ODS — 72 / 66

Strengths (verified): bounded worksheet graph (repeat-run logical addressing,
merges, coverage, bounded writers); transactional owners for metadata,
calculation settings, annotations, protection, DataPilot, and embedded charts
(but named-range and RDF edits are `&mut self` facade mutations;
protection/DataPilot/charts lack Patch types); the ODF 1.4 tracked-changes
owner (now 7,606 lines) is the family's best code — all four record classes
(`table:insertion`, `table:deletion`, `table:movement`,
`table:cell-content-change` per part-3 §9.9.3/§9.9.9/§9.9.13/§9.9.17, with
dependencies/cut-offs/nested deletions also modeled), limits, reversible
patches, no-op byte preservation, verified against real LibreOffice files;
flat `.fods` has a complete `FlatSpreadsheet` owner
(Snapshot/Transaction/Patch/Commit, `src/flat.rs`, 713 lines, active tests);
DDE sources and scenarios are wired and read-only-exposed; **the orphaned
~6,200-line `src/codec/content/` tree is deleted and the blanket
`allow(dead_code)` is gone** (S3); **conditional formatting (1,192 lines) and
sparklines (899) are no longer fully unwired** — they gained `parse()`
validators and are inventoried per-sheet through the new bounded, inert
`source_features` module (722 lines, own `SourceFeatureLimits`,
DOCTYPE/entity rejection) with two active test files
(`tests/source_features.rs`, `tests/source_feature_inventory.rs`); **204
tests green**.

Deductions (verified): still no typed CF/sparkline attach/edit surface — the
inventory is read-only, not feature CRUD; exactly 12 `&mut self` facade
mutations and the `MutableSpreadsheet` attached root against
ADR-0003/0023 (S4); only three `Patch` structs crate-wide (annotations,
flat, tracked-changes) — DataPilot/protection/charts have none; the
inventory tests are thinner than the deferred suites they replace; the ODS
`tests/deferred/README.md` contains a stale "`FlatSpreadsheet` no longer
exists" line.

### ODP — 65 / 61

Strengths (verified): strong read model (**22** `DrawingShapeKind` variants
across all `draw:*` shape elements and all six `dr3d:*` elements — an
earlier revision said 23, off by one; animation `Kind` covers the complete
ODF 1.4 part-3 §15 SMIL vocabulary — exactly the 12 `anim:*` elements);
dedicated transition/action/settings/page-layout models, 1,214-line
layout-master codec; a from-scratch Builder (1,414 lines — grew with
media-plugin serialization); localized edits for masters/handout/
annotations/RDF; the slide/shape editor is real — the 802-line `pub(super)`
`MutablePresentation` backs the public `authoring::edit::Snapshot` API with
16 green tests; **the 9 red doctests are fixed — 21/21 doctests pass** (S1);
`draw:plugin` frame-children now have a full inert model with XLink
validation, schema-validated parsing, and Builder serialization
(`model/media.rs`, parser tests `codec/parser/tests.rs:684-718`); new public
re-export modules (`image`/`master`/`slide`, `FlatPresentation`, shared
`odf-common` authoring types) and a `Presentation::snapshot()` entry point;
real-file tests hard-fail on error; 167 unit/integration tests green.

Deductions (verified): **the crate is now clippy-red itself** — 46 lib / 47
lib-test errors under CI flags (`let_underscore_must_use`, wildcard-match,
…), and its dependency `litchi-odf-common` adds 1,084 errors (S1) — a new
regression introduced by the fix commits; chart transactions require a
manual `commit().into_bytes()` → `from_bytes` republish (deliberate but
still friction); `Shape` is a pub-field struct with a type tag and
`Option<String>` coordinates — exactly the pattern ADR-0004 forbids;
encrypted open exists via shared machinery but **no in-crate test exercises
an encrypted ODP**; no forms, encryption write, or signatures;
`draw:applet` and `draw:floating-frame` remain unmodeled; a stale internal
comment at `model/animation.rs:15` still cites "ODF 1.3, Part 3, section
10".

### ODG / ODI / ODB, ODC, ODF formula, OTH / ODM

`d1b775d4` grew these shells substantially; none is a shell anymore:

- **ODG — 40 / 60** (was 22 / 40): the semantic model is now wired —
  `package/snapshot.rs` (822 lines) parses pages/layers/shapes into the model
  types on open (new `model/shape.rs`: `ShapeKind`,
  `Shape::{name, layer, kind, text, frame}`), and the facade exposes
  `pages()/layers()/edit()` with a package-level
  `Transaction::set_shape_text` → `Commit` → `Patch{is_applicable_to,
  apply, inverse}` (`facade/mod.rs:45-68,146-345`), tested by
  `tests/package_semantics.rs` (3 tests). The flat chain and real-content
  fixture stand. Editing is still single-verb (shape text only). 1,815 src
  lines, 13 tests, zero production panic sites.
- **ODI — 35 / 50** (was 17 / 30): a full flat ADR-0003 chain now exists —
  `FlatImageTransaction::{set_frame_name, set_source}` (`flat.rs:87,121`,
  812 lines) → `FlatImageCommit` → `FlatImagePatch::{is_applicable_to,
  apply, changes, inverse}`, re-exported and tested (`tests/flat_edit.rs`).
  Package facade is read-only; still no real-file fixtures. 1,117 src lines,
  11 tests.
- **ODB — 30 / 45** (was 15 / 30): typed content parsing is back (S3) —
  bounded, inert, limit-guarded schema + query catalog
  (`model/catalog.rs`, 600 lines) tested against real LibreOffice `.odb`
  files (`tests/schema_catalog.rs`, 6 tests); `codec/content.rs` is a
  189-line NsReader structural validator. Forms/reports/connections remain
  unmodeled; everything is read-only. 1,173 src lines, 14 tests.
- **ODC — 60 / 72** (was 48 / 55): 1,022 src lines; chart parsed on open;
  flat axis edits with the full in-memory patch chain (`flat.rs:127,265,280`
  — note: no `is_applicable_to` in this crate); **the chart class is now the
  typed §19.15 vocabulary** from `litchi-odf-common::chart::class` (S4/ODT
  section), facade `Chart::class() -> Result<ChartClass>`. The real
  capability lives in `litchi-odf-common::chart` (now 2,889 lines: 1,632
  authoring + 619 reader + 638 view/glue incl. the 235-line class module).
  Mutation surface is still axis-edit-only.
- **ODF formula — 68 / 72** (unchanged; untouched by both commits): small
  but complete — a **34-kind** validated MathML tree (33 MathML kinds +
  `Other`) with byte-exact round-trip, `.odf`/`.otf` packages, and atomic
  `set_root` (serialize → re-validate → swap, failure leaves the original
  untouched, `facade/mod.rs:87-95`).
- **OTH — 28 / 45** (was 15 / 38): 681 src lines; the substring check is
  gone — `codec/content.rs` (325 lines) does `compact_xml` + NsReader
  structural validation of the `office:document-content/office:body/
  office:text` envelope with DTD/entity/depth rejection; semantic types are
  wired via facade `text_body()` → `codec::paragraphs` projection; a
  source-checked no-op `Edit/Commit/Patch` with `inverse()` exists
  (`facade/mod.rs:107-170`) — correct transaction shape, zero actual
  mutation capability. Still no real-file tests.
- **ODM — 35 / 50** (was 20 / 40): 746 src lines; real parsing now —
  `codec::parse` yields `Semantics` (title + ordered subdocument references;
  `model/subdocument.rs`, 127 lines, `Reference`/`Target::is_external`),
  facade `Master::{title, subdocuments}`, tested against a real LibreOffice
  master-document fixture (`tests/semantic_references.rs`, 3 tests). The
  NsReader structural pass with DOCTYPE rejection/depth caps stands. Still
  no editing of master documents.

(`litchi-odraw` is not an ODF crate at all — it is the MS OfficeArt binary
record library, and its clippy gate is now green per S1.)

### iWork (Pages 63 / Keynote 64 / Numbers 63; API 71 / 71 / 70 — unchanged)

Neither fix commit touched any iWork crate code; every claim was re-verified
at HEAD anyway.

Strengths (verified): **the stack compiles** (S1); clean layered
architecture — 17 crates under ADR-0028 (iwa-monolith-exit), with the legacy
structured host seam retired; complete physical layer (ZIP/Snappy/exactly 40
protobuf schemas + 9 hand-written projections); an extraordinarily wide
editor write surface — **PagesEditor: exactly 704 pub fns across 85 impl
files (not 98 as previously stated) including 186 chart editors;
KeynoteEditor: ~690 pub fns with full build animations (typed `Effect`
model, bounded `Unknown` fallback for unmodeled Apple effects);
NumbersEditor: ~655 pub fns with broad conditional-highlight predicates**
(no duration predicates; row/column inserts on volatile-highlight tables are
*rejected*, not translated); wire-level patching that preserves unknown
protobuf fields byte-for-byte, with targeted tests; source-checked package
apply (fingerprint + byte + semantic re-check) and byte passthrough when
unmodified; ADR-0003-shaped reversible transactions — Pages section/body
text (`SectionTextEdit` with `SectionTextPatch::inverse()`,
`package/section_text.rs`) and Keynote speaker notes (`SlideNotesEdit`, a
1,464-line module); native Apple-authored fixtures with pinned SHA-256s back
integration tests in all three crates (5/5 pass at HEAD); excellent panic
posture (exactly one guarded unwrap family-wide — though see S5 for the 50
expects / 29 `unreachable!()` / one `panic!` the "panic-free" phrasing
omits — all re-verified exact).

Deductions (verified): **the wide editor surface lives in legacy
`litchi-iwa` and is unreachable through the semantic crates** —
`litchi-pages`/`litchi-numbers`/`litchi-keynote` do not re-export the
editors at all; only the newer, narrower transaction APIs are; the
format-neutral `litchi::iwork` facade is read-only (but
`litchi::pages`/`litchi::keynote` wholesale re-exports do expose package
write transactions at top level); Pages body text is editable only as
plain-text splices within single-storage sections (rich formatting,
footnotes, inline objects, section breaks out of scope); atomic temp+rename
save exists only in legacy `litchi-iwa` — the semantic crates have no
save-to-disk; **Numbers has no package transaction family**; no tracked
changes; no encrypted-document support (typed refusal, `.iwpv2` marker →
`ErrorKind::Encrypted`); no formula evaluation (AST only; `litchi-eval` is
not wired); legacy editors remain `&mut self` clone-staging with no
reversible patches; no per-crate FEATURE_MATRIX. ~~ADR-0023
double-numbering~~ FIXED — renamed ADR-0029 and indexed (S8).

### Markdown — 55 / 64

Export-only by design (verified: no Markdown parser anywhere in the
workspace). Nearly all remaining first-draft deductions were closed in
`d1b775d4`/`46e88ebc`: headings resolve from DOCX `style_id`, DOC/RTF
`outline_level`, **and now ODT outline levels** (`markdown_heading_levels`
sidecar, `crates/litchi/src/document/doc.rs:260-307`, consumed at
`markdown/document.rs:143`); list detection uses `numbering.xml` semantics
for DOCX/DOC with the textual heuristic kept only as a fallback for
RTF/ODT/Pages; body text is escaped; the parallel path propagates errors;
formula conversion failures return `Err`; **underline is implemented**
(`write_underlined`, inline HTML, `writer.rs:752-810`, dispatched in
`write_run`); **the silent-drop problem is converted to typed errors** —
`Document::validate_markdown_projection` (`crates/litchi/src/document/
doc.rs:158-258`, wired at `markdown/document.rs:128`) makes DOC/DOCX/RTF/ODT
hyperlinks, footnotes, DOC/DOCX/ODT inline images, and RTF quote fields
return `Err(Unsupported)` instead of being silently dropped; the dead-code
leftovers are purged; the emission layer (now 2,489 lines in
`crates/litchi/src/markdown/`) plus the leaf crate hold 29 tests combined
(12 emission + 17 leaf), plus 2 integration tests and the first real-fixture
test (`testPictures.doc` inline-image refusal); ADR-0024:492-501 documents
the crate and blesses the centralized-adapter topology.

Remaining deductions (verified): hyperlinks/images/footnotes are still not
*converted* — only refused loudly; blockquotes (outside RTF quote fields)
and code blocks remain silently dropped; Pages paragraphs get no heading
detection (`writer.rs:384`); no golden-output corpus — validation is
synthetic in-memory packages plus the one real-fixture refusal test. The
options vocabulary is typed and the leaf crate is panic-free (zero
unwrap/expect/unreachable/panic in `litchi-markdown/src/`).

---

## Overall assessment

First tier — DOCX, XLSX, PPTX, RTF: deep functionality, solid test corpora,
broadly conformant APIs, and — new in this revision — **all four build and
lint clean**, with RTF's 965-error backlog and PPT's 7,675 clippy errors
fully paid. The common gap remains a unified editing model for existing
documents per ADR-0003 (XLSX excepted). Second tier — XLS/XLSB/DOC/PPT:
strong readers and create-only writers; open→edit→save for main content is
missing across the board (S6), but both silent-data-loss traps are closed
(mutations now fail loudly), DOC gained bounded opens and a zeroizing
`Password`, and PPT gained a real transition writer. XLS's clippy gate is
still red, now blocked by the newly surfaced `litchi-ole-common` (1,062
errors). ODF: the ADR-0023 split's regressions are now essentially
adjudicated — **ODT's packaged root has a complete, tested transaction
chain** (S4 resolved, conformance 52 → 70), the ODS orphan tree and blanket
`allow(dead_code)` are gone, ODP's doctests are green (though its own clippy
newly went red), and **none of the small families is a shell anymore**: ODB
parses typed schema catalogs from real files, ODG wires its semantic model
with a package-level reversible edit, ODI has a full flat patch chain, ODM
parses real master-document references, OTH validates structurally, and ODC
gained the typed §19.15 chart class. ODI/OTH remain the weakest (no
real-file tests; OTH's only patch is a no-op) but the "should not be
advertised" verdict now applies to OTH alone. iWork: unchanged since the
baseline — compiles, native-fixture-backed, with ADR-0003-shaped text/notes
transactions, but the vast editor surface is still reachable only through the
legacy crate. Markdown: the silent-drop integrity problem is fixed (typed
refusals), though conversion coverage is still thin. The next-largest
workspace-level debts, in order: the S6 binary open→edit→save hole; the
red clippy layer (`litchi-ole-common`, `litchi-odf-common`, `litchi-odp`);
the ODT/ODF-common panic-macro backlog; and the still-absent ADR-0003 wire
patch ecosystem everywhere.
