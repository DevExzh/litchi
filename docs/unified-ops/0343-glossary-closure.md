# 0343 — ADR 0034 glossary closure and docs hygiene (audit WP14)

- Status: implemented on `uo2/glossary-closure`, cut from
  `feat/unified-ops-api` at `ed832736c3`; trunk merged at `d03195dfea`
  (record 0341) and again after records 0340 and 0342 merged, before the
  final gates.
- Date: 2026-10-18.
- Input: the definition-of-done audit of 2026-10-17, WP14: the pending 🔴
  cells of `docs/API_GLOSSARY.md`, the records index of `PROGRESS.md`
  (it stopped at 0248), and stale facade, README and feature-matrix docs.
- Owned: the 🔴 public names of the leaf crates, `docs/API_GLOSSARY.md`,
  `tools/api_glossary.json`, the records index of `PROGRESS.md`, the facade
  and `litchi-ops` crate docs, `README.md` and `docs/FEATURE_MATRIX.md`.

## 1. Method

Every 🔴 cell was checked against the public API, not against the
glossary text. Rustdoc JSON (all features) was built for `litchi-core`,
`litchi-opc`, `litchi-ooxml-common`, the OOXML, legacy, ODF and iWork format
crates, `litchi-ops` and the facade, and walked with the gate's own
`Surface` class (`tools/check_api_glossary.py`) to list every exported path
that still spells a banned name. Counts below are exported paths from that
walk (re-exports count once per path).

Several 🔴 cells were stale: the work had landed in an earlier record but
the cell was never flipped, and a merge had left duplicate **Rollback**,
**No-op introspection**, **Mutable writer** and **Metadata read/mutation**
rows separated by blank lines (outside the table). The duplicates are
merged into one row each, keeping every ✅ citation.

The status legend gains ⏸: deferred by a numbered decision of this record
(D1–D11 in §4) that the owner memo lists. After this record the glossary has
no 🔴 cell.

## 2. Renames implemented (option a)

Each is a hard removal with every workspace caller migrated (format crates,
`litchi-ops`, the facade, its tests and examples, `tools/perf-baseline`);
the old names are added to the gate's ban lists where the gate can express
them.

| Glossary row | Change | Commit | Gate |
|---|---|---|---|
| Stream output | `OpcPackage::to_stream`, `PackageWriter::write_to_stream`, `SourceArtifact::write_to_stream{,_with_accounting}`, `SourcePartSplicePlan::write_to_stream{,_with_accounting,_with_expected_artifact}` → `write_to*` (72 files) | `2f7bdddd5c` | `to_stream`, `write_to_stream*` banned in `opc-minimal`/`opc-full` |
| Patch base mismatch | DOCX `document::Error::StaleSource` and the source-backed `story_text`, `paragraph_copy`, `tail_append` `StaleSource` variants → `PatchConflict`, as PPTX did in 0252. `litchi-ops` keeps the `io_error` classification | `628572c84b` | `StaleSource` banned in `docx-*` |
| Transaction type, Rollback | The 13 XLSB scoped families still named `Transaction` (`calculation_chain`, `comments::threaded`, `custom_data`, `data_model`, `external_link`, `connections`, `web_extension_bindings`, `shared_workbook`, `slicer`, `theme`, `timeline`, `volatile_dependencies`, `xml_maps`) → `Edit`; the consuming `web_extension_bindings::Edit::rollback` and the root alias `CustomDataTransaction` are deleted (no caller used either) | `acb42936c2` | `Transaction`, `rollback`, `CustomDataTransaction` banned crate-wide in `xlsb-*` |
| Open from bytes | The flat owners `odt::flat::Document`, `ods::FlatSpreadsheet`, `odg::FlatDrawing`, `odc::FlatChart`, `odp::FlatPresentation` and the flat ODP `Snapshot`: `from_bytes(Vec<u8>)` → `from_bytes(&[u8])` (the bytes are copied, as every package owner does; the `*_with_limits` variants keep owned input) | `2dea26322c` | signature only; verified by compilation |
| Positional selection word | `DocumentView::{paragraph, paragraph_text}`, `PresentationView::slide`, the PPTX handle's `slide`, compat `Table::row_at` and `Row::{cell_at, grid_span_at}`: `usize` → `impl Into<litchi_core::Position>` (`From<usize>` keeps call sites unchanged). A rustdoc-JSON scan of `litchi-ops` finds no other positional `usize` parameter | `032e892b13` | signature only |

The gate checks names, not signatures, so the two signature rows rely on
compilation and the rustdoc scan above.

## 3. Rows flipped by verification (no code change)

- **Transaction type / Rollback / No-op (scoped PPTX):** `Transaction`,
  `rollback` and `is_changed` have been banned crate-wide in the `pptx`
  profiles since 0252; the walk finds no exported PPTX `Transaction` or
  `rollback`.
- **Metadata read / mutation (DOCX):** `Package::metadata()`,
  `edit_metadata` and `apply_metadata` exist and `props`/`props_mut`/
  `put_props`/`clear_props` are banned (0251). The two stale 🔴 duplicate
  rows are deleted; the XLSB source-owner note they carried moves into the
  surviving row.
- **Option-bearing fresh authoring:** DOCX `Builder` (0251), PPTX
  `builder::Builder` (0127), DOC `Builder` (0188); XLSX has no
  option-bearing authoring and creates through `Package::new()` (0250).
  `WorkbookWriter`/`Writer` are banned in xlsb/doc/xls/ppt. Only the
  `litchi-iwa` prefixes remain (D11).
- **Semantic limits:** W1 is complete (0297–0299); `DocumentLimits` and
  `DocumentSemanticLimits` are banned in core, numbers and keynote and no
  format crate exports them. `xml_minifier::audit::DocumentLimits` is the
  minifier's XML-audit profile, not a document semantic limit.
- **iWork metadata write (Keynote, Numbers):** a root metadata `set` is a
  typed refusal (`unsupported_capability` for Keynote,
  `unsupported_property` for Numbers) and leaves the bytes unchanged — the
  canonical "typed unsupported until native gates pass". New test
  `crates/litchi-ops/tests/iwork_unified.rs::iwork_metadata_write`.

## 4. Deferrals (option b)

Each is added to the owner memo (0331) as items 14–24 (item 21, D8, is
withdrawn); the adopted default is in effect until the owner overrides it.

- **D1 — Direct package mutators (row *Immutable owner snapshot*).**
  `litchi_docx::Package` exposes 85 public `&mut self` methods (VBA, fonts,
  custom XML, glossary, mail merge, task panes, signing, `custom_props_mut`,
  the `save*`/`to_*` outputs …), `litchi_pptx::Package` 77 (`add_slide_master`,
  `put_*`/`remove_*`, `apply_*_patch`, outputs) and XLSB keeps
  `Package::into_workbook` with the mutating `Workbook` verbs (0261 F1).
  Moving each into an `Edit` family (or making the outputs `&self`) is a
  per-verb redesign of the two crates, not a rename. Default: kept; no new public `&mut self` mutator may be
  added on a `Package`.
- **D2 — Scoped OOXML commit triad (rows *New snapshot accessor*,
  *Diagnostics*, *Split commit*, *Patch application*, *No-op*).** About 200
  scoped commits in DOCX (45 `snapshot` paths, 18 `into_snapshot`), PPTX
  (50), XLSX (92) and XLSB (24) return a part-level `Snapshot`, not a
  `Package`; most have no `Diagnostics`, and DOCX (19 paths) and XLSX (30)
  still name scoped staging types `Transaction`. Renaming `snapshot()` to
  `package()` would misname a part snapshot; the canonical shape needs each
  family to publish a `Package` and carry `Diagnostics`, which changes their
  semantics. Default: the primary transactions stay canonical; scoped
  families keep their shape until each is migrated with its owner.
- **D3 — PPT scoped families (row *Rollback*, W4 residue).** 12
  `Transaction` families and 19 `rollback` methods (`font`,
  `picture_bullets`, `client_anchor`, `client_data`, `diagram`, `broadcast`,
  …). `rollback` there restores staged record state used by the PPT
  writers; removing it needs the writers re-plumbed. Default: kept as W4
  residue.
- **D4 — `UnsafeEdit` refusal vocabulary and ODT `MutableDocument` (row
  *Mutable writer*).** `Error::UnsafeEdit` is the general typed edit refusal
  in DOCX (about 40 call sites) and PPTX (root and `opened`); renaming it
  means choosing its replacement taxonomy (`Unsupported`, `PatchConflict`,
  `Refused { reason }`), an error-API decision. ODT `MutableDocument`
  (0268 G2) is the authoring model behind the ODT `Builder`. Default: kept.
- **D5 — PPT limits-only ingress (row *Configured open*).** `raw::Reader`
  and 173 record- and family-level `*_with_limits` entry points take
  `SourceLimits` directly. Folding them into `OpenOptions` is a W4 design
  (record parsers take limits, not open options). Default: kept.
- **D6 — Flat ODF and ODC limits (row *Coverage policy*).**
  `FlatDrawing` and `FlatPresentation` take no caller-selected limits (fixed
  ceilings); the flat ODT/ODS/ODC owners take their crate `Limits`, not
  `OpenOptions`; ODC keeps its chart `Limits`. Adding an options envelope to
  the flat owners is new API whose shape (one ODF flat `OpenOptions` or one
  per crate) is the ODF owner's call. Default: the fixed ceilings stay.
- **D7 — `Metadata.keywords` type (row *Semantic metadata type*).** Owner
  decision D2 (2026-10-15) unblocked the iWork callers, but `Metadata`
  derives `Serialize`/`Deserialize`: the field type fixes the metadata
  shape of JSON views, the CLI and `litchi-py` dictionaries (a joined
  string today). Default: `Option<String>` stays, the OOXML codec keeps
  keyword mixed content losslessly.
- **D8 — DOCX app/custom properties through the shared engine (row
  *Metadata mutation*).** Withdrawn: record 0342, merged while this record
  was in flight, added `metadata::Edit::{app_props, set_app_props,
  custom_props, set_custom_props}` over the shared engine's part patch and
  `Package::app_props()`, so the cell is ✅ (0342). The direct mutator
  `Package::custom_props_mut` stays under D1.
- **D9 — Content redaction call-site audit (OOXML/legacy).** An audit of
  error and `Debug` strings across four OOXML and three legacy crates, not
  a rename. Default: the rule stands; violations are fixed as found.
- **D10 — iWork grid coordinates (row *Grid coordinates*).**
  `CellPosition`/`CellRange` → `At`/`Area` touches about 1,300 occurrences
  in 161 files of the iWork crates, `litchi-ops` and the facade, including
  `format/iwork/keynote`, which record 0340 was editing while this record ran; `litchi-ops`
  already has a private `format::iwork::tree::At`, and several ops files
  import the XLSX `At`/`Area`. Default: deferred to the next iWork wave
  (0340 has since merged, so it is unblocked).
- **D11 — `litchi-iwa` builder prefixes (row *Option-bearing fresh
  authoring*).** `PagesDocumentBuilder`, `KeynoteDocumentBuilder` and
  `NumbersDocumentBuilder` (about 940 occurrences in 183 files) live in the
  unpublished host crate `litchi-iwa`, which no workspace crate depends on
  and no glossary profile covers; the leaf crates already have canonical
  `Builder`s (0297–0299). Whether to rename inside a crate the unified API no
  longer uses, or to delete the crate, is the owner's call. Default: kept.

## 5. Docs hygiene

- `PROGRESS.md` records index: rows 0249–0343 added (one per record file,
  0251a–0251f included), each with its title and the merge that landed it,
  taken from the record's first lines and its trunk merge subject.
- Facade crate docs (`crates/litchi/src/lib.rs`, *Unified operations*):
  "Batches and the other formats' unified write handlers are still in
  progress" replaced by the current state (every format reads and writes
  through `File`) with a compiled doctest of `create`, `add` and `set`.
  The `litchi-ops` crate *Status* section is corrected the same way.
- `README.md`: the "does not do yet" section (editing, iWork, other ODF
  formats) and the *Current Status* list are rewritten; the removed
  `Package::presentation_mut` and `*_with_limits` snippets are replaced
  (`builder::Builder`, `open_with_options`).
- `docs/FEATURE_MATRIX.md`: the stale "format-neutral read-only iWork
  facade" and "facade authoring is not implied" statements are corrected; a
  *Unified API* section and the missing RTF row are added.

## 6. Verification

After the last trunk merge (`6994d038de`, records 0340–0342):

| Gate | Result |
|---|---|
| `cargo fmt --all --check` | 0 |
| clippy `-D warnings`, all targets, all features: opc, docx, xlsb, odt, odg, odc, odp, ppt, pptx, xlsx, ooxml-common, ops, litchi, litchi-py; `litchi-ods --lib` | 0 |
| clippy `litchi-ops --no-default-features` alone and with `docx`, `pptx`, `xlsx`, `xlsb`, `ppt`, `doc`, `rtf`, `odt`, `ods`, `odp`, `odg`, `odc`, `keynote`, `numbers` | 0 each |
| `RUSTC_BOOTSTRAP=1 tools/check_api_glossary.py` | 0: 40 profiles, 2,329 banned names, 253 families |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features` (changed crates, ops, litchi) | 0 |
| tests, all features: docx 2,455; xlsb 1,295; odt 1,128; odg 162; odc 59; odp 444; ods 1,774 (every target but the purged-fixture `ods_formula_conditional_native`); ooxml-common 513; xlsx 2,467; pptx 1,842; ppt 1,499; opc | all pass |
| `cargo test -p litchi-ops --all-features` in chunks (lib 948, doc 215, 162 test targets in 7 chunks) and `cargo test -p litchi --all-features` (235) | all pass |
| capability/registry goldens, all features and all-but-`encryption` | pass unchanged (no re-bless needed) |
| fuzz (`CARGO_TARGET_DIR=target/fuzz`): `litchi-core`, `litchi-ops`, `litchi-ops --features apply` checks; both `smoke_batch_payload` release runs; `crates/litchi/fuzz` check | 0 each |
| `tools/perf-baseline`: `cargo check --locked --all-targets` | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 (293 tests) |
| `tools/check_feature_matrix.py` | 0 (31 ok) |
| tool unit tests, capability-delta/corpus/bench-matrix `--check`, crate boundaries, non-iWork gate | 0 |

`crates/litchi/fuzz` (the facade's iWork targets) did not compile before
this record: three targets named `litchi_pages::OpenOptions`, which that
workspace does not depend on (now `litchi::pages::OpenOptions`), one used a
removed `cell::comment::reply` path (now `cell::comment::transaction`), and
one had an ambiguous `as_ref()`. None of the three was caused by this
record's renames; they are fixed here.

## 7. Performance

No hot path changed shape. Two calls gained a copy: the flat ODP open in
`litchi-ops` (`FlatPresentation::from_bytes(&bytes)`) and the default-limit
flat `from_bytes` constructors copy their input once before parsing, which
is linear in the XML that the parser then scans anyway. No measurement was
taken; the unified flat ODP open is not a benchmarked route.
