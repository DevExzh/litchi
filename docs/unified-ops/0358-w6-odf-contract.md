# 0358 — W6: ODF on the record 0351 contract

- Status: implemented on `uo2/w6-odf`, cut from `feat/unified-ops-api` at
  `4ff05a303e`, merged with trunk twice (`f7bd3399c7`, `f997a1ed8c` at trunk
  `7893492f9e`, which carries records 0355, 0355a, 0356 and 0357).
- Date: 2026-10-20.
- Input: owner decision [0350](0350-owner-api-debt-decision.md); contract
  [0351](0351-api-contract.md) §2, §4.3, §5 (wave W6); the source-backed
  family shape of [0355a](0355a-w1-docx-gaps.md) §1.
- Owned: `crates/litchi-{odf-common,odt,ods,odp,odg,odc,odi,odm,oth,odb,odf-formula}/**`;
  the ODF hunks of `crates/litchi-ops/src/{error.rs, error/tests.rs,
  format/{odt,ods,odp,odf_*}*}` and of `crates/litchi-ops/examples/*`;
  `crates/litchi-odf/**`; `crates/litchi/examples/odt_*`;
  `tools/perf-baseline`, `tools/native-resave`; the ODF profiles of
  `tools/api_glossary.json`; the W6 line of `docs/API_GLOSSARY.md` §2a, its
  ODF semantic-dimension table and coverage row;
  `test-data/odf/formula-conditional-native/`.
- `performance_claim`: measured (§8).

## 1. Summary

The eleven ODF crates now conform to record 0351 §2:

- every `Package`, read-only projection and flat owner is immutable,
  including ODT `generic::Package`, `generic::flat::Document` and
  `odc::Document`; the gate's `immutable_owners` lists them all;
- ODT `MutableDocument` (136 `&mut self` verbs) and its
  `Content`/`ContentMut`/`Styles`/`StylesMut` views are crate-private:
  `Builder` authors, `Package::edit()` edits, and gained every structural
  verb the model had (§3);
- every commit publishes the format's owner with `Patch` and `Diagnostics`
  (`Package::apply` returns a `Commit` everywhere); part-level commits and
  patches are crate-private (§4);
- no public `Transaction`, `rollback`, `is_changed`, `into_snapshot`,
  `Mutable*` or `*_with_limits` name remains in any ODF crate; the gate bans
  them crate-wide in all eleven profiles, with `Limits` in every `flat`
  module and the deleted prefixed flat names (§7);
- flat owners live at `flat::<Projection>` and open with `OpenOptions`; the
  flat and ODC chart limits fold into `SourceLimits`/`SemanticLimits` (§5);
- module errors carry `Refused(Refusal)`, `PatchConflict` and `Signed`, and
  the typed refusal enums expose `kind()` (§6).

Every ODF round-trip, CRUD, corpus and no-op test passes; the OfficeCLI
corpus, the capability and registry goldens and the batch-equivalence tests
are unchanged, so no golden was re-blessed. Consumers were migrated in the
same commits. One gap remains (§9): the four source-backed stream
publishers keep their stream shape.

## 2. Commits

| Commit | Scope |
|---|---|
| `a2a0220ceb` | odf-common foundation: `*_with_options` constructors, limits required on `detect::*`, `compact_xml::validate`, `validate_package`; `Error::{Refused, PatchConflict, Signed}` and their ops classification; metadata families return a `Commit`; `Package::apply` returns a `Commit` in all ten format crates; 19 `SemanticLimitKind` dimensions |
| `c400b06268` | ODM `title`/`link` families; ODB/ODG `PreparedEdit` → `SubEdit`; `kind()` on the ODB/ODG refusal enums |
| `48071820a2` | flat owners `odg::flat::Drawing`, `odc::flat::Chart`, `odi::flat::Image`, `odp::flat::Presentation`, `ods::flat::Workbook`, `odt::flat::Document` with `OpenOptions` and six-name families; ODC `Limits` folded; ODC definition drafts |
| `d34966847d` | ODS part families behind `Package::edit`, `builder::Workbook`, required formula-parser limits |
| `ff774ea123` | ODS: the purged `ods_formula_conditional_native` fixture restored as a tracked compact fixture |
| `757ab35c20` | ODT: immutable generic and chart owners; `MutableDocument` behind `Package::edit`; `header_footer` and `protection` families crate-private; renames |
| `116286bf63` | ODP: crate-private chart drafts; `tail_append` module |
| `613b69de2c` | gate registrations and glossary |
| this record | record, ABBA evidence, `odf_w6_probe`, bench-matrix parser |

## 3. ODT

- `mutable` is crate-private. `Package::edit()` gained 121 verbs
  (`transaction/document_verbs.rs`): headings, lists, tables, sections
  (`add`, `wrap`, `unwrap`, `update`, `remove`, `clear`), notes
  configurations, line numbering, bibliography configuration, font faces,
  outline, list and drawing styles, page sequence, master pages and page
  layouts, header/footer text, XML and properties, bookmarks, reference
  marks, text indexes and marks, tracked changes, RDFa and `text:meta`,
  XForms, variable declarations, form properties and the twelve control
  kinds, images and text boxes. Each stages an `Operation::Document` step;
  consecutive steps share one working `MutableDocument` at commit, and the
  result passes the same compact-XML audit as every other operation. A
  parity test (`tests/internal/document_verbs.rs`) checks a verb chain
  byte for byte against the model and its inverse. The steps have no
  durable wire form (like the structural body edits); `OperationResult::Name`
  reports an allocated frame name. The deprecated `update_paragraph` alias
  has no verb (`replace_paragraph` stays).
- `Document` gained the reads only the model had: `bookmark_targets`,
  `form_properties` and the twelve `*_controls` lists.
- `generic::Package` (12 mutators), `generic::flat::Document` (13, was
  `generic::FlatDocument`) and `odc::Document` (11) are immutable. Each has
  `edit() -> Edit`, `apply(&Patch) -> Commit` and the six-name family
  (`generic`, `generic::flat`, `odc`, one `staged_family!` macro); `Edit`
  verbs stage the former mutators, `Patch` holds the exact before/after
  bytes (`is_applicable_to`, `is_noop`, `inverse`). The flat document opens
  with `open_with_options`, `from_reader_with_options`,
  `from_bytes_with_options` and the `*_with_execution_context(.., options,
  context)` forms (the `u64` `*_with_limit` forms are crate-private);
  `with_execution_context(&self, context)` returns a re-charged copy and
  leaves the source and its reservation untouched.
- `header_footer`: the layout `Transaction`, `Commit`, `Patch`,
  `DurablePatch`, `History`, `MergePlan`, `Resolution` and `Target` are
  crate-private; `Snapshot` stays the read view with
  `prepare_master_page_transfer`/`prepare_section_layout_transfer`. New
  verbs `insert_section_layout`, `replace_section_layout`,
  `remove_section_layout`, `insert_master_page_transfer(transfer, available)`
  and `insert_section_layout_transfer(transfer, available)` keep every
  layout capability on the package route. The layout-level durable patch,
  history and three-way merge are deleted: the package `Patch` replays and
  inverts these edits and `JoinedEdits` composes package edits.
- `protection`: the policy `Transaction` family (constructible only inside
  the crate) is crate-private; its durable patch, transfer and merge, which
  no route reached, are deleted.
- Names: `JoinedEdit`/`joined_edit` → `JoinedEdits`/`joined_edits`;
  `body::{TextPropertyEdit, ParagraphPropertyEdit, TabStopEdit}` →
  `TextPropertyUpdate`, `ParagraphPropertyUpdate`, `TabStopSpec` (ODP
  `scene::TextPropertyEdit` likewise); `EmbeddedResourceGcRefusal::kind()`.

## 4. ODS and ODP

- ODS part families (option A'): the closure-scoped drafts stay public as
  the closure arguments of `document::Edit` and `builder::Workbook`
  (`annotations`, `content_validation`, `protection` and `tracked_changes`
  `Transaction` → `Edit`; `rollback` → `reset`; a draft's `snapshot()` →
  `current()`). Part `Commit`, `Patch`, `update`, the catalog `Transaction`s,
  the `metadata`/`settings` XML transactions and `Snapshot::edit` are
  crate-private; their exact-source replay/inverse code, which only the
  removed routes used, is deleted. Read `Snapshot`s stay public.
- New `document::Edit` routes so no capability loses its package path:
  `database_ranges`, `scenarios`, `table_templates`, `dde`, `sheet_metadata`
  and `set_settings`; `apply_content_validation_patch` is gone (the package
  patch covers it).
- `MutableSpreadsheet` → `builder::Workbook` (0351 §2.1 rule 7): the same
  setters and closure editors, `build()` publishing a `Package` instead of
  `to_bytes`/`into_package`/`save`, no `apply_*_patch`; the builder's
  database-range draft is `database_range::Draft`.
- Limits: `Expression::parse(source, &Limits)`, `FormulaParser::parse(&limits)`
  and `Reference::parse`/`reference::parse(value, &Limits)` take required
  limits; the tracked-change and content-validation `*_with_limits` twins are
  crate-private (the public forms apply the default limits, the precedent of
  0357 §4). `advanced::LogicalRowEdit` → `LogicalRowOperation`;
  `EncryptionWriteRefusal::kind()`.
- ODP: the chart `Edit`/`Commit`/`Patch`/`Diagnostics` and the inventory
  `Transaction`/`Editor` are crate-private (`Package::edit()` owns every chart
  verb); the unused inventory transaction is deleted. The tail-append types
  are `tail_append::{Edit, PublicationPlan, Error, Limits, SourceProof}`;
  `CryptoRefusal::kind()`.
- Tests that drove now crate-private APIs moved into the crates as unit
  tests (`tests/internal/`, included by `src/internal_tests.rs`: 19 ODS and
  24 ODT suites); only their assertions on deleted part patches were
  removed. `ods_formula_conditional_native` reads
  `test-data/odf/formula-conditional-native/observations.json` (345 KiB, a
  lossless compact re-encoding of the purged 5.5 MiB evidence, with
  provenance), so the full litchi-ods suite builds and runs again.

## 5. Limits and flat owners (§2.5)

- Flat owners `flat::{Document, Workbook, Presentation, Drawing, Chart,
  Image}` (and ODT `generic::flat::Document`) open with `open`,
  `from_bytes`, `from_reader` and their `*_with_options(.., OpenOptions)`
  forms, keep the options (`options()`), and replay with `apply(&Patch)`.
  The bound of a flat document is the smaller of the source `input_bytes`
  and the semantic `content_bytes`. ODC `Builder::with_options`,
  `Package::options`, `Patch::from_bytes(bytes, OpenOptions)`.
- `SemanticLimits` has 19 dimensions, registered in the glossary (§4, "ODF
  semantic dimensions"). Each default is the most permissive bound an ODF
  reader enforced, so flat documents now open under the shared ceilings
  (flat ODT 16 → 64 MiB, flat ODP 64 → 256 MiB, nesting 4,096 where readers
  had 256 or 512).

## 6. Errors (§2.4)

`litchi_odf_common::Error` gained `Refused(Refusal)`, `PatchConflict
{ operation, reason }` and `Signed`; stale patches, foreign patches and
staged plans made on another base are `PatchConflict`, a changed write on a
signed package is `Signed`, encrypted metadata writes are
`Refused(UnsupportedOperation)`. litchi-ops classifies them by kind
(`invalid_input`, `document_protected`). `kind()` on ODB `ProtectionRefusal`,
`ComponentTransferRefusal`, ODG `SecurityLifecycleRefusal`, ODS
`EncryptionWriteRefusal`, ODT `EmbeddedResourceGcRefusal`, ODP
`CryptoRefusal`.

## 7. Gate (§4.3)

`tools/api_glossary.json`, all eleven ODF profiles: crate-wide bans
`rollback`, `Transaction` (exact, suffix), `is_changed`, `into_snapshot`,
`Mutable` (prefix), `_with_limits` (suffix); `Limits` in every `flat`
module; the deleted prefixed flat names at the crate roots (and
`FlatDocument` at `litchi_odt::generic`); families `flat` (six crates),
`litchi_odt::{generic, generic::flat, odc}`, `litchi_odm::{title, link}`
next to the primary and metadata families; immutable owners as in §1.
Stale scopes of deleted types are dropped or moved. The debt inventory
(`tools/api_debt_inventory.py`) of the ODF crates went from 490 rows to 71:
41 are the `ods::builder::Workbook` draft the inventory's name heuristic
reads as an owner (0351 §2.1 rule 7 names it), 19 are `Commit` aliases of
`litchi_core::edit::Commit` it cannot see through (the gate checks them in
`litchi-core`), 2 the `litchi_core::SubEdit` composition aliases of ODB and
ODG, and 9 the source-backed families of §9.

## 8. Performance

`crates/litchi-ops/examples/odf_w6_probe.rs` times the unified routes the
wave touched with only the `litchi_ops` API, so one source builds on the
base (trunk `7893492f9e`) and the head; release, pinned to one CPU, 61
rounds per scenario, run A-B-B-A (`results/0358/abba.sh`,
`analyze.py`, `summary.json`; binary hashes in `binaries.sha256`).
Ratio = median of the after runs / median of the before runs.

| Route | Before (µs) | After (µs) | After/before |
|---|---|---|---|
| ODT open + get (1k paragraphs) | 138.9 | 141.8 | 1.021 |
| ODT set | 418.3 | 415.4 | 0.993 |
| ODT batch, 1k sets | 1790.8 | 1806.1 | 1.009 |
| ODS open + get (100×10) | 1280.8 | 1293.4 | 1.010 |
| ODS set | 11.5 | 11.4 | 0.995 |
| ODS batch, 1k sets | 5066.6 | 5129.2 | 1.012 |
| ODP open + get (20 slides) | 120.2 | 121.0 | 1.007 |
| ODP set | 424.2 | 425.2 | 1.002 |
| ODG open + get | 2315.5 | 2295.4 | 0.991 |
| ODG set | 9965.0 | 10070.6 | 1.011 |
| ODC open + get | 78.0 | 77.1 | 0.988 |
| ODC set | 205.0 | 201.6 | 0.984 |
| flat ODT open + get (1k paragraphs) | 372.0 | 381.8 | 1.026 |
| flat ODS open + get (100×10) | 2153.3 | 2210.5 | 1.027 |

Every route is within the 1.05 regression budget. The largest changes are
the flat opens (ODT +2.6 %, ODS +2.7 %; two earlier A-B-B-A runs gave +0.3–1.4 %
and +3.0–3.4 %). The flat owners now admit their input through
`OpenOptions` at each open; the difference was not profiled further because
it is inside the budget. `tools/unified_bench_matrix.py`
gains the `p0358` parser, one after/before row per route.

## 9. Gaps and follow-ups

1. **Source-backed stream publishers** keep their shape: ODS
   `SourceCellEdit`/`SourceCellCommit`/`SourceCellPatch`,
   `dde::Source{Edit, Commit, Patch}`,
   `sheet_metadata::SourceMetadata{Edit, Commit, Patch}` and ODP
   `tail_append` publish to a writer; their commits do not yet return a
   committed `SourceBackedSpreadsheet`/`SourceBackedPresentation`. Moving
   them onto the deferred-package shape of record 0355a §1 is the follow-up
   (the §2a line keeps it red, like W2's `SourceBacked*` editors).
2. Deleted part-level machinery, which no public route reached after the
   wave: ODS part patches, the ODT layout-level durable patch, history and
   merge, the ODT protection durable patch/transfer/merge, the ODP chart
   inventory transaction. Their package-level counterparts (`Patch`,
   `inverse`, `JoinedEdits`, `History`) remain.
3. `litchi-ops --no-default-features --features odf-objects` (the internal
   umbrella with no object format) does not compile, on trunk `7893492f9e`
   as on this branch; it is not a format feature, and every format feature
   alone passes clippy.
4. A few common-core `Error::Unsupported` sites in the ODF crates still
   carry refusals as strings rather than `Refused(Refusal)`.

## 10. Gates

All on the merged head, one heavy command at a time
(`CARGO_BUILD_JOBS=6`, `RUST_TEST_THREADS=8`).

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-{odf-common,odt,odp,odg,odc,odi,odm,oth,odb,odf-formula,odf} -p litchi-ops -p litchi -p litchi-py -p litchi-cli --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ods --all-features --all-targets -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features odt\|ods\|odp\|odg\|odc\|odi\|odm\|oth\|odb\|odf-formula] -- -D warnings` (none and each) | 0 |
| `cargo test --no-fail-fast --all-features -p` each ODF crate and `litchi-odf` (litchi-ods `--all-targets`, every target including `ods_formula_conditional_native`), plus `--doc` | 0 (4,521 passed) |
| `cargo test -p litchi-ops --all-features` (lib, every integration test in chunks, doc) and `cargo test -p litchi --all-features` | 0 (2,850 passed) |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p` the ODF crates | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` | 0 (44 profiles, 3,038 banned names, 419 families, 67 immutable owners) |
| `python3 tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py verify` | 0; 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets`; `cd tools/native-resave && cargo check --locked --all-features`; the three ODF fuzz crates `cargo check --offline --all-targets` | 0; 0; 0 |
| `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references tools.test_check_crate_boundaries tools.test_non_iwork_gate tools.test_check_api_glossary` | 0 |
| `python3 tools/officecli_capability_delta.py --check`; `python3 tools/officecli_corpus.py --check`; `python3 tools/unified_bench_matrix.py --check`; `python3 tools/check_feature_matrix.py` | 0; 0; 0; 0 |

The capability and registry goldens and the OfficeCLI corpus pass without
re-blessing: no ODF capability, code or output changed.
