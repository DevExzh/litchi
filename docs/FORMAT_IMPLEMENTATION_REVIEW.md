# Format Implementation Review — `feat/office-format-completeness`

Date: 2026-08-08
Scope: every format family in the workspace, reviewed against the source code,
the normative specs under `3rdparty/specs/`, and the ADRs under `docs/adr/`.
Method: per-format deep audits that read the actual code, counted
`unwrap`/`expect`/`panic!` in production paths, verified FEATURE_MATRIX claims
against the implementation, and ran `cargo test` / `cargo clippy` where the
environment allowed. Documentation claims were treated as unverified until
confirmed in code.

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
   merge — is implemented **nowhere**. Every format stops at in-memory
   reversible patches; only XLSX completes the full in-memory chain
   (snapshot → transaction → commit → source-checked apply → inverse patch).
2. Rendering, pagination, and field evaluation are explicitly declared
   out of scope in every FEATURE_MATRIX and are not penalized.

## Score summary

| Format | Functional | API conformance | Position |
|---|---|---|---|
| DOCX | 85 | 76 | One of the two most complete crates |
| XLSX | 80 | 84 | Best ADR-0003 conformance; matrix overclaims |
| PPTX | 78 | 70 | Broadest part coverage; opened files not editable |
| RTF | 76 | 66 | Strong read/write; lossy, no transactions, lint red |
| DOC | 73 | 68 | Near-full read spectrum; main content not editable |
| XLSB | 71 | 76 | Read+write; matrix contains false claims |
| XLS | 70 | 70 | Wide read, create-only writer; test gate currently red |
| PPT | 71 | 62 | Strong read+create; narrow edit surface; lint badly red |
| ODT | 78 | 52 | Deepest ODF crate; frontally violates ADR-0003/0006 |
| ODS | 64 | 60 | Upper-middle; the family split caused regressions |
| ODP | 57 | 58 | Read-strong, write-weak; key editor is dead code |
| Numbers | 63 | 68 | Wide editor; branch currently does not compile |
| Keynote | 60 | 68 | Same compile break |
| Pages | 58 | 68 | Body story not editable; compile break |
| ODF formula | 68 | 72 | Small but self-consistent |
| ODC | 40 | 45 | Usable only via `litchi-odf-common::chart` |
| Markdown | 33 | 52 | Early export-only prototype, zero tests on the emitter |
| ODG / ODI / ODB | 15 | 30 | Validated package shells; not implemented formats |
| OTH / ODM | 15 | 38 | Thin shells; far from ADR-0023 family slices |

---

## Severe findings (cross-format)

These are the issues that most urgently undermine the branch's support claims,
ordered by severity.

### S1. Build and quality gates are red on the committed branch

ADR-0008 requires continuously buildable phases and warning-denied lint gates.
As committed, none of the following pass:

- **iWork family does not compile.** `crates/litchi-iwa-archive/src/limits.rs:225`
  constructs `soapberry_zip::office::ArchiveLimits` with 3 fields; commit
  `816979b4` added `max_compressed_size`, `max_member_name_bytes`,
  `max_metadata_bytes` and updated `litchi-opc` but missed `litchi-iwa-archive`.
  The entire `litchi-iwa*` stack, `litchi-pages/numbers/keynote`, and the
  facade `iwork` feature fail with E0063; ~1,900 iWork tests are unrunnable.
- **`litchi-ppt` clippy: ~5,204 errors** on stable 1.97.1 (including 80
  `expect_used` and 51 `unwrap_used`, both workspace-denied lints with zero
  in-crate allows). CI runs `cargo clippy --workspace --all-targets
  --all-features -- -D warnings` without `continue-on-error`.
- **`litchi-rtf`: 963 lint errors** (mostly `unused_qualifications`);
  ADR-0008 itself records this as a pre-existing backlog, still unpaid.
- **`litchi-xls`: 3 committed integration tests fail deterministically**
  (`tests/xls_external_tables.rs:215,317` — XML map ↔ list-object XML source
  integration never closed; `tests/xls_row_block_index.rs:65` — test requests
  sheet index 13 from a 13-sheet fixture), and clippy fails on the
  `litchi-odraw` dependency with 160 errors.
- **`pyo3-litchi` and the facade `office`/`all` aggregate features do not
  compile** (rtf lint errors + iwa-archive E0063, both in committed code).
- **`tools/check_crate_boundaries.py` reports 5 unclassified edges**
  (`litchi-crypto→litchi-core`, `litchi-ooxml-common→litchi-crypto`,
  `litchi-ppt→litchi-fonts`, `litchi-pptx→litchi-crypto`,
  `litchi-xlsx→litchi-crypto`) — the boundary JSON was not updated when commit
  `23035db5` added the crypto dependencies.
- **`crates/litchi/tests/encryption_facade.rs` lacks its
  `#![cfg(feature = "encryption")]` header**, so `cargo test --tests` fails to
  compile without the encryption feature.

### S2. FEATURE_MATRIX overclaims proven false by code

- `litchi-xlsb/docs/FEATURE_MATRIX.md:92` claims password encryption
  ✅ Read / ✅ Write. The crate has no `encryption` feature, does not depend
  on `litchi-crypto`, and encrypted XLSB is a CFB container that cannot pass
  the ZIP/OPC open path. **The claim is false.**
- `litchi-xlsx` matrix claims ✅ Write for conditional formatting (the module
  itself states it has no package layer — read-only), ✅/✅ for hyperlinks
  (raw passthrough only), ✅ for defined-name writes (`workbook/model.rs:470`
  is an inert read-only record), typed page breaks (absent), and
  core/extended properties on the facade (requires escaping to
  `into_plain_opc()`). Five rows claim more than the code does.
- `litchi-xlsb` slicer/timeline rows say "transactional inert CRUD"; the
  actual API is part-level `load/store` functions with no
  Snapshot/Transaction/Commit/Patch.
- `litchi-ods` rows for DDE sources and scenarios are marked 🟡, but
  `src/model/dde.rs` and `src/model/scenario.rs` are **dead files not wired
  into the module tree** — the rows describe code that does not compile.
- `litchi-markdown/src/lib.rs` claims per-format `impl ToMarkdown` blocks
  "live alongside their respective format crates"; all implementations are in
  fact centralized in `crates/litchi/src/markdown/`.

### S3. The ODF family split (ADR-0023) caused silent feature regressions

The split moved previously-tested capability into `tests/deferred/`
directories that Cargo never compiles (no `[[test]]` targets; the files
reference deleted pre-split APIs such as `litchi_odf::DrawingDocument`):

- **97 deferred test files in `litchi-odt`/`litchi-odf-common`**, including
  all cross-family corpus round-trips (`odf_corpus.rs`), parser-hardening
  suites (17 files), and flat-format read/write (22 files). The "lossless
  preservation" and "malformed-input robustness" claims for ODF currently
  have **no in-build verification**.
- `litchi-ods`: 17 deferred files (conditional formatting, hyperlinks, in-table
  shapes/images, sparklines, DDE, tracked-change authoring) plus 4 dead source
  files (`codec/evaluation.rs`, `model/dde.rs`, `model/scenario.rs`,
  `model/rich_text.rs`).
- `litchi-odp`: the 1,591-line `MutablePresentation` (full slide/shape CRUD)
  is dead code — `authoring/mod.rs` declares only `pub mod builder` — while
  doc comments in `package/presentation.rs` still reference it.
- ODG drawing-style resources and ODB schema/query support existed pre-split
  and are now gone; `test-data/odf/corpus/` (15 real files) and
  `test-data/odf/drawing/` (7 flat drawings) have no active consumer.

### S4. ADR-0003 is frontally violated by the ODF text family

`litchi-odt::Document` exposes ~25 `&mut self` methods that mutate the package
directly (`document/package.rs:213-549`); there is no Edit/Commit/Patch model
anywhere in the family, and `MutableDocument` is precisely the attached
mutable root that ADR-0023 step 5 says must be deleted once replacements are
verified. Structural edits take raw `usize` indices and return `Result<()>`
instead of selector-first `Result<Option<_>>`. The same `&mut self` pattern
recurs in `litchi-ods` and `litchi-odp` facades.

### S5. Panic-free discipline (ADR-0006) is unevenly enforced

Production-path `unwrap`/`expect` counts after stripping `#[cfg(test)]`:

- `litchi-odt`: **424**; `litchi-odf-common`: **90** — direct conflict with
  the workspace `unwrap_used = "deny"` / `expect_used = "deny"` lints.
- `litchi-ppt`: 131 (51 unwrap + 80 expect), zero allows.
- `litchi-xlsx`: 151 guarded unwraps; `litchi-docx`: 28 unwrap + 34 expect;
  `litchi-xlsb`: 13 unwrap + dozens of invariant expects + ~10
  `unreachable!()`; `litchi-doc`: ~75; `litchi-imgconv`: ~30 (some reachable
  from untrusted EMF path state).
- `litchi-odt` lacks `#![forbid(unsafe_code)]`; `litchi-opc` also lacks it.
- Concrete ADR-0006 violation: `MutableComment::new` stamps
  `chrono::Utc::now()` (`litchi-docx/src/writer/comment.rs:35`) — new files
  consult ambient time despite the determinism requirement.

### S6. No general open→edit→save path for existing binary Office files

For XLS, PPT, and DOC the writer is create-only and the reader is immutable;
editing an existing file is possible only through a handful of per-feature
transaction modules (XLS: 8; PPT: 20; DOC: 13+). There is no public path to
change a cell value in an existing `.xls`, rewrite a text box in an existing
`.ppt`, or edit a paragraph in an existing `.doc`. Worse, `litchi-ppt`'s
`shapes::Shape` exposes public mutators (`set_text`, `shapes/shape.rs:377`)
with **no serialization path at all** — mutations are silently lost. This is
the single largest functional gap against the "full Office CRUD" goal of
ADR-0001, and it is shared by PPTX (opened presentations' slide content is
locked behind `UnsafeEdit`, enforced by `tests/edit_guards.rs`).

### S7. Zero real-file verification for iWork

`test-data/` contains no `.pages`/`.numbers`/`.key` files, and no test
references the libetonyek QA samples present under `3rdparty/libreoffice-core`
(e.g. `Pages_4.pages`, `Keynote_1-6.key`). Every iWork round-trip is
self-generated → self-read; compatibility with actual Apple output has no
regression evidence. The same "self-generated only" caveat applies to
`litchi-crypto` (no real encrypted Office files), `litchi-sign` (keys
generated on the fly), `litchi-vba` (no real `.docm`/`.xlsm` fixtures), and
`litchi-drawingml` / `litchi-spreadsheet-drawing` / `litchi-ograph` (no real
files at all — coverage is only indirect via host formats).

### S8. Governance and documentation gaps

- `docs/CRUD_Scenario_Checklist.md` is referenced by `docs/adr/README.md:11`
  and ADR-0008 as a support gate but **does not exist**.
- ADR-0023 is double-numbered (`0023-odf-family-crate-split` and
  `0023-iwa-index-foundation`); the README indexes only the former.
- `litchi-markdown` is absent from ADR-0024's topology and from every ADR;
  its support claim has no audit baseline.
- ODG/ODC/ODI/ODB/OTH/ODM, the ODF formula crate, and the three iWork crates
  have no FEATURE_MATRIX documents at all.
- `docs/FEATURE_MATRIX.md:44-45` still says the shared ODF implementation
  "lives primarily in `litchi-odf`" — that crate is now a 48-line umbrella.
- `crates/litchi-xls/README.md` demonstrates a nonexistent `XlsWorkbook`
  type (the export is `Workbook`).
- `docs/report/odf-iwa-rtf.md` cites pre-split paths that no longer exist.

---

## Per-format details

### DOCX — 85 / 76

Strengths: ~114k lines; all [MS-DOCX] extension families 2.1–2.13 have read
models except the umalqura calendar (2.2.7); modern comments, conflict
revisions, SDT checksums, OpenType extensions each have 1,000+-line codecs;
full package layer (OPC, Strict/Transitional, MCE, signatures, encryption,
read limits); writer covers paragraphs/tables/sections/revisions/comments/
fields/SDT/OLE/VML/SmartArt/watermarks/TOC. 900+ tests green, real POI and
LibreOffice fixtures, a fuzz target, and recorded macOS Word open/inspect
evidence.

Deductions: no document-level ADR-0003 model — body editing is
`&mut MutableDocument`, and the 13 part-level transactions have no unified
patch type; `Block` lacks the ADR-0007 `Unknown` fallback and there is no
public `Inline` enum (`document/model.rs:69-76`); `litchi-word` (which holds
the `Visibility` projection vocabulary) is an orphan crate no one depends on,
so visible/review/all projections are unimplemented; no `litchi-math` crate
(math is raw OMML, `src/math.rs`); ambient-time violation
(`writer/comment.rs:35`); raw relationship IDs leak into the public API
(`hyperlink.rs:88`, `package/merge.rs:19`) against ADR-0004; 28 unwrap + 34
expect in production (all invariant-guarded but nonzero); SDT
appearance/checkbox/repeating-section semantics are partial.

### XLSX — 80 / 84

Strengths: the most complete ADR-0003 implementation in the workspace —
immutable `Arc` snapshot, `edit()` → transaction → `Commit{workbook, patch}`
→ `Patch::inverse()` → source-checked `Workbook::apply()`, structured
`ConflictSet/JoinError` (`edit/model.rs:1003-1565`); ~30 part-level owners
(calc chain per ADR-0018, data validation, query tables, connections,
slicers, timelines, rich values, a 5,667-line pivot module); ADR-0017
producer templates fully compliant with byte-parity tests; zero panic macros
in production; 810 tests with 73 real fixtures from three corpora.

Deductions: the five FEATURE_MATRIX overclaims listed in S2; Survey parts
(spec 2.1.9) entirely absent and undocumented; no ChartEx, dynamic-array
spill, or formula evaluation (honestly marked); patches have no JSON wire
form/`seal()`/`History` (self-acknowledged debt at `edit/model.rs:1401`);
`NumberFormat{pub id, pub code}` accepts anything unvalidated
(`style/stylesheet/number_format.rs:11-24`) and `CellFont.color` is a raw hex
`String`, violating ADR-0004/0007; 151 guarded production unwraps.

### PPTX — 78 / 70

Strengths: broadest part coverage in the workspace — every [MS-PPTX]
extension family 2.2.1–2.2.20 has a bounded typed reader; ~15 feature domains
with high-quality reversible transactions (notes, fonts per ADR-0022, table
styles per ADR-0020, sections, custom shows, guides, designer tags,
comments, OLE/ActiveX, 3D models, zooms); ADR-0004 exemplars `shape::Scene`
and the `time::Offset` decimal-time grammar; 587 tests green, 76 real
`.pptx` files, 13 save→reopen loops.

Deductions: **opened presentations are not editable** — `presentation_mut()`
returns `UnsafeEdit` for opened packages (`package/model.rs:63-66`, locked by
`tests/edit_guards.rs`); slide add/remove/reorder and shape editing work only
on newly created packages; **ADR-0013 is unfulfilled** — the promised
package-level `notes()`/`put_notes()`/`remove_notes(slide)` API does not
exist and note deletion requires drilling to physical PackURIs
(`tests/pptx_notes_crud.rs:253-267`); `litchi-slide` defines the ADR-0007
`LayoutRole`/`Review`/`Look` vocabulary but nothing depends on it (dead
code); chart/table creation is not on the facade; the main facade is
`&mut self` + whole-graph clone rollback rather than the snapshot model; no
pptx-specific fuzz target.

### XLS (BIFF8) — 70 / 70

Strengths: ~112k lines; wide typed read coverage (SST, full styles, pivots,
chart metadata, three encryption profiles, signatures, VBA metadata,
revisions, toolbars, XML maps); substantial create-only writer with 46
write→reopen test files; ADR-0012 (checked formula references) and ADR-0027
(sheet anchors) fully implemented; 1,165 tests with 37 real fixtures
including encrypted files.

Deductions: **3 committed tests fail deterministically and clippy is red**
(S1); typed record coverage ~209/466 (~45%), with opaque records not
guaranteed to survive typed mutation (honestly documented); no general
open→edit→save for existing workbooks; write-side formulas are a ~30-variant
Ptg subset (no PtgName/FuncVar/Ref3d/array constants); ADR-0016's own
migration debt unpaid — public writer APIs still take raw integers
(`set_column_width(col: u16)`, `freeze_panes(u32)`, `set_row_height(u32)`);
`Worksheet` has public mutators against ADR-0003; credentials are `Clone`
and weak XOR encryption can be written without an explicit policy, both
against ADR-0006.

### XLSB — 71 / 76

Strengths: read+write; 539/875 record kinds named (~62%); 568 tests green;
12 real fixtures including producer-quirk regression tests; 9 feature domains
with source-checked Snapshot/Transaction/Commit/Patch; consistent laziness
for external content.

Deductions: the false encryption claim (S2); semantic coverage roughly half
the spec (Data Model, rich values, MDX, smart tags, ActiveX absent — mostly
declared out of scope); no cell/style transaction editing on opened
workbooks; writer is create-only; no whole-book preserve re-save; 13
production unwraps plus invariant expects/`unreachable!()`; the package-level
`Error` is not `#[non_exhaustive]`; the unified facade exposes XLSB
read-only.

### PPT (binary) — 71 / 62

Strengths: 100% RecordType recognition (218/218); ~75 read accessors on
`Presentation`; a create-from-scratch writer (shapes, rich text, tables,
pictures, animations, transitions, sounds, VBA, encryption); 20 ADR-0003
transaction modules; 1,081 tests green on default features with 25 real
fixtures and byte-exact round-trip assertions; CryptoAPI encryption read and
write verified against real POI files.

Deductions: no public path for general content edits of existing files;
**`shapes::Shape` public mutators have no serialization path — silent data
loss** (`shapes/shape.rs:377`); `Writer::add_chart` unconditionally returns
`UnsupportedAuthoring`; dual-track shapes API (`Box<dyn Shape>` trait objects
alongside the `ShapeEnum` data enum) against ADR-0004; ~300 top-level
re-exports with long names; 131 production unwrap/expect against denied
lints and ~5,204 clippy errors (S1); `--all-features` tests could not be run
in the review environment.

### DOC (binary) — 73 / 68

Strengths: near-full read spectrum — FIB, piece table, FKP/bin tables,
styles, numbering, all seven stories, dozens of typed fields, comments,
bookmarks, revisions, sections, tables, plus long-tail structures (mail
merge, smart tags, command bars, route slip, MTEF equations, OLE controls);
a 19.7k-line full-stack create writer; 13+ per-feature transaction editors
including a story-level `RevisionEditor`; 1,081 tests green, 40 real
fixtures plus POI encrypted files, one fuzz target; panic-free record is
good (0 todo/unimplemented).

Deductions: no unified snapshot/edit/patch model for main content — the
crate is "read-only reader + create-only writer + per-feature patchers";
patches are binary-stream replacements with no JSON form or reversibility
grading; **no memory budgets on the main read path** (`Document::from_ole`
loads whole streams; `Package::open` has no limits parameter) against
ADR-0005; no accepted/rejected text projections (ADR-0007); password passed
as plain `Option<&str>`, no `Locked`/`Sensitive` type-state (ADR-0006);
DOP 2007/2010/2013 and DopMth unmodelled; one unreproduced test flake
observed; the FEATURE_MATRIX implementation map cites pre-refactor paths.

### RTF — 76 / 66

Strengths: very wide read coverage (~400 control-word variants plus an
`Unknown` fallback, nested tables, shapes, legacy drawing, OLE objects,
dozens of typed fields, EQ+OMML math, revisions, compressed RTF both
directions per [MS-OXRTFCP]); the immutable snapshot facade is a textbook
ADR-0003/0004 implementation (cheap shared `Arc` handle, lazy derived values,
compile-fail doctests proving immutability); 996 tests, 42 real corpus files,
and prefix-truncation / byte-mutation sweeps proving no panics.

Deductions: **unknown destinations/control words are dropped at parse time
and writes are lossy** (`skip_group`, with tests asserting the dropping) —
against ADR-0006's preserve-by-default rule and the crate's own "lossless
snapshot" claim; **no Edit/Commit/Patch model at all**, and `raw::Document`
exposes bare `&mut` setters on an attached tree (against ADR-0001's
"mutation is tracked"); 963-lint backlog (S1); no dedicated FEATURE_MATRIX;
the writer `indent` option is dead code; Mac code pages 10001/10007 are
rejected (an intentional, typed, defensible choice).

### ODT (+ common/umbrella) — 78 / 52

Strengths: deepest ODF implementation (~111k lines) — full package
lifecycle, encryption (AES/Blowfish/Argon2id), signatures verified against a
real LibreOffice XAdES signature, RDF, manifest; complete text structures
(paragraphs/tables/lists/bookmarks/tracked changes/sections/fields/indexes/
ruby/forms/page sequences); flat `.fodt`; 807 tests green with real
`.fodt` and LibreOffice QA files; detection API fully matches ADR-0009 with
fuzz targets.

Deductions: see S3 (deferred tests) and S4 (ADR-0003 violation); 424 + 90
production unwrap/expect (S5); `usize`-index selectors returning `Result<()>`;
missing `#![forbid(unsafe_code)]`; migration leftovers (dead directories
`odt/src/ods/`, `odf-common/src/migration/`, an unreferenced master-document
fixture, an `include!`-based module); ADR-0009's "sole owner" text not
updated after detection moved to `litchi-odf-common`.

### ODS — 64 / 60

Strengths: bounded worksheet graph (repeat-run logical addressing, merges,
coverage); transactional owners for named ranges, metadata, calculation
settings, annotations, protection, DataPilot, embedded charts, RDF; the new
ODF 1.4 tracked-changes owner (7,335 lines) is the family's best code — all
four record classes, limits, reversible patches, no-op byte preservation,
verified against real LibreOffice files (`change-tracking.ods`,
`RecordChangesProtected.ods`); 165 tests green.

Deductions: conditional formatting / sparklines / DDE / scenarios are dead
or unwired files (making two matrix 🟡 rows untrue); flat `.fods` cannot be
opened through the facade (detection only); no style cascade; `&mut self`
facade edits and the `MutableSpreadsheet` attached root against
ADR-0003/0023; deferred tests hide the split regressions (S3).

### ODP — 57 / 58

Strengths: strong read model (shape kinds across `draw:*`/`dr3d:*`, SMIL
animations, transitions, actions, settings, page layouts, masters); a
from-scratch Builder; localized edits for masters/handout/annotations/RDF;
encrypted open; 123 tests green.

Deductions: the slide/shape editor for existing documents is 1,591 lines of
**dead code** (`authoring/mutable.rs` not declared in the module tree) while
doc comments still reference it; chart transactions require a manual
`commit().into_bytes()` → `from_bytes` republish; the only real-file test
silently skips on parse failure (`let Ok(..) else { return }`); `Shape` is a
pub-field struct with a type tag and a string of `Option`s — exactly the
pattern ADR-0004 forbids — with `Option<String>` coordinates; no forms,
encryption write, or signatures.

### ODG / ODI / ODB — 15 / 30; ODC — 40 / 45; ODF formula — 68 / 72; OTH / ODM — 15 / 38

ODG, ODI, and ODB are isomorphic package shells (MIME + body-marker
validation + raw XML snapshot + content.xml-replacing Builder). Their public
model types (`Layer`/`Page`, `Frame`/`Source`, `Connection`/`Query`) are
disconnected from any codec — they can neither be parsed from nor written to
a real document — and there is no edit/transaction surface. They should not
be counted as implemented formats. ODC is a 336-line shell whose real
capability lives in `litchi-odf-common::chart` (2,598-line authoring +
582-line reader), but chart class is a free `String` rather than the part-3
§11 typed vocabulary. `litchi-odf-formula` is small but complete: a ~40-kind
validated MathML tree with byte-exact round-trip, `.odf`/`.otf` packages, and
atomic `set_root`. OTH/ODM are ~260–290-line snapshot shells with
substring-level validation (`xml.contains("<office:text")`), unwired semantic
types, and no real-file tests — far from ADR-0023's "complete family slice".
(`litchi-odraw` is not an ODF crate at all — it is the MS OfficeArt binary
record library, and is in good shape on its own terms.)

### iWork (Pages 58 / Numbers 63 / Keynote 60, API 68 each)

Strengths: clean three-layer architecture; complete physical layer
(ZIP/Snappy/40 protobuf schemas); an extraordinarily wide editor write
surface (PagesEditor: 184 pub fns including ~60 chart editors;
KeynoteEditor: 224 pub fns with full build animations; NumbersEditor: full
conditional-highlight predicates with dependency translation); wire-level
patching that preserves unknown protobuf fields byte-for-byte, with targeted
tests; package-level ADR-0003 fully compliant (source-checked apply, atomic
save, byte passthrough when unmodified); excellent panic-free posture (one
guarded unwrap family-wide).

Deductions: **the branch does not compile** (S1); **zero real iWork files in
any test** (S7); the facade exposes the three semantic crates read-only, so
iWork has no write path at the top level unlike peer formats; Pages body
story is not editable (text boxes are richly editable, the main text flow is
not); no tracked changes, no encrypted-document support (typed refusal), no
formula evaluation (AST only); editors are `&mut self` with internal
clone-staging and publish no reversible semantic patches; no per-crate
FEATURE_MATRIX; the ADR index's 0023 double-numbering leaves the
iwa-index-foundation ADR unindexed.

### Markdown — 33 / 52

Export-only by design (no parser anywhere in the workspace). The core
semantic mapping is broken: **headings are unsupported** — `write_paragraph`
never reads paragraph styles, so Heading 1–6 become plain paragraphs;
hyperlinks, images, blockquotes, code blocks, footnotes, and underline are
all unimplemented and silently dropped; list detection is textual heuristic
rather than `numbering.xml` semantics, so auto-numbered lists are missed;
body text does not escape Markdown special characters. **The 2,030-line
emission layer in `crates/litchi/src/markdown/` has zero tests** and no
real-file/golden validation. The options vocabulary itself is typed and the
leaf crate is panic-free, but parallel paths swallow errors
(`document.rs:45,54`), formula failures embed placeholder strings, and two
`unreachable!()` sit in hot paths. The crate is absent from ADR-0024 and
every ADR — a governance blind spot — and its lib.rs architecture claim
(per-format impls in format crates) does not match reality (centralized in
the umbrella).

---

## Overall assessment

First tier — DOCX, XLSX, PPTX, RTF: deep functionality, solid tests, broadly
conformant APIs; the common gap is a unified editing model for existing
documents per ADR-0003. Second tier — XLS/XLSB/DOC/PPT: strong readers and
create-only writers, but open→edit→save is largely missing and several
quality gates are red. ODF: the ADR-0023 split is unfinished and has caused
regressions now hidden in `tests/deferred/`; ODT is functionally deep but
frontally violates the transactional and panic-free ADRs. iWork: an
impressively wide write surface with no real-file backing and a currently
broken build. Markdown and the small ODF families (ODG/ODI/ODB/OTH/ODM)
should not be advertised as supported in their current state.
