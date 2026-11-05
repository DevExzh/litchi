# 0358a — W6 review: source-backed ODF families, no lost capability

- Status: implemented on `uo2/w6-odf` after record
  [0358](0358-w6-odf-contract.md), merged with trunk twice: `7f36f7b909`
  (`81e4e8705c`, records 0356a and 0357a) and `ff0f1fa036` (`129c3861a4`,
  the 0358 merge and records 0356b and 0357b).
- Date: 2026-10-20.
- Input: the coordinator's review of 0358 (five items, the owner's
  no-exceptions/no-loss direction); the source-backed family shape of
  [0355a](0355a-w1-docx-gaps.md) §1 and its PPTX twin
  [0356a](0356a-pptx-review-follow-ups.md) (`source::Package`,
  `source::svg`); the durable/merge/transfer reference of
  [0357a](0357a-ppt-record-sized-patches.md).
- Owned: as 0358, plus `crates/litchi-ops/examples/odf_source_backed_probe.rs`
  and `docs/unified-ops/results/0358a/`.
- `performance_claim`: measured (§6); every changed route writes
  byte-identical output.

## 1. Summary

| Review item | Done |
|---|---|
| 1. Source-backed publishers on the 0355a §1 shape | ODS `source_backed::{cells, dde, sheet_metadata}` and ODP `tail_append` are registered families; `edit_x()` on `SourceBackedSpreadsheet`/`SourceBackedPresentation` → `Edit` → `commit()` publishing a deferred owner with `Patch` and `Diagnostics`; both owners gained `write_to`/`write_to_with_options`/`save`/`save_with_durability` (§2) |
| 2. Every deleted part-level capability covered | each item with its replacement route and the test that proves it (§3); the ODT layout and protection durable patch, history, merge and transfer, the ODS part patches and the ODP chart inventory transaction all have a package-level durable, joined/three-way and transfer route |
| 3. `--features odf-objects` alone compiles | `cfg(odf_objects)` from a build script; the feature matrix covers it (§4) |
| 4. No string-carrying `Unsupported` | every one is `Refused(Refusal)` with a static, content-free reason, or `Error::Other` for internal failures (§4) |
| 5. Inventory heuristic for builder drafts | left to the PPTX wave (0356b), as directed |

## 2. Source-backed families (item 1)

```text
SourceBackedSpreadsheet::edit_cells()                 -> source_backed::cells::Edit
SourceBackedSpreadsheet::edit_dde() / edit_dde_with(limits, &context)
                                                      -> source_backed::dde::Edit
SourceBackedSpreadsheet::edit_sheet_metadata() / edit_sheet_metadata_with(..)
                                                      -> source_backed::sheet_metadata::Edit
SourceBackedPresentation::edit_tail_append(title, body) -> tail_append::Edit
x::Edit::{verbs..}(&mut self, ..)                     // staging only
x::Edit::commit(self) -> Result<x::Commit, x::Error>  // validates, writes nothing
x::Commit = litchi_core::edit::Commit<Owner, x::Patch, x::Diagnostics>
Owner::apply_x(&self, &x::Patch) -> Result<x::Commit, x::Error>
Owner::write_to(&self, impl Write) / save(path)        // the only output
```

**ODS.** The three families change only `content.xml`. A commit runs every
check the former `commit` ran (row-local eligibility and readback for cells;
the DDE and metadata renderers, their limits and execution context, and the
signed-source refusal) and returns a new `SourceBackedSpreadsheet` sharing
the retained source (`Arc<SourceBackedPackage>`) and holding the candidate
`content.xml`; nothing is written. `write_to` streams the source through the
same `write_content_xml_to_stream_with_options` the former
`SourceCellCommit::write_to`/`dde::SourceCommit::write_to`/`SourceMetadataCommit::write_to`
called, so the bytes are identical (§6); an unedited spreadsheet writes its
source exactly. Reads of a committed spreadsheet (`sheets`, `cell`,
`cell_batch`, `text`, `settings`, `definitions`, `scenarios`, `dde`,
`sheet_metadata`, `member_data("content.xml")`, `materialize`, the tolerance
report) parse its `content.xml` once (`effective()`); memory stays bounded by
`content.xml`, never the archive, and a caller that only saves never pays the
parse. Chained edits work (`commit.package().edit_cells()`), and the cached
row layout stays keyed to the exact `content.xml` it was scanned from.

- `Patch` holds the exact before and after `content.xml` (`is_noop`,
  `inverse`, `source_xml`, `target_xml`). It replays onto a spreadsheet over
  the same retained source whose `content.xml` is exactly its before content
  and refuses any other with `Error::PatchConflict`; its inverse applied to
  the commit's spreadsheet restores the source content, and writing that
  writes the exact source archive (tested in all three families).
- `apply_dde_with`/`apply_sheet_metadata_with(patch, limits, &context)` keep
  the former destination-snapshot admission: the target's output ceiling, an
  output reservation and a full parse under the destination policy.
- `Diagnostics`: `cells::Diagnostics { changed(), changed_cells() }`;
  `dde`/`sheet_metadata` use `litchi_core::edit::Diagnostics`. `Error` and
  `LimitKind` are `litchi_odf_common::{Error, LimitKind}` (the ODS package
  family's), so a refusal is `Error::Refused`, a stale patch
  `Error::PatchConflict`.
- Reads: `dde()`/`dde_with()` return `source_backed::dde::Snapshot`,
  `sheet_metadata()`/`sheet_metadata_with()` return
  `source_backed::sheet_metadata::Snapshot` (the former `SourceSnapshot` and
  `SourceMetadataSnapshot`, without `edit()`).
- `write_to`/`write_to_with_options` return the typed
  `SourceContentPublicationReport`/`SourceContentPublicationError` (sink
  progress kept); `save` stages and renames atomically
  (`litchi_core::atomic`). `litchi_odf_common::Error` gained
  `From<SourceContentPublicationError>` and `From<SourceContentInsertionError>`.

**ODP.** `tail_append::Edit` holds only the title, body, `Limits`
(`set_limits`) and `SourceContentPublicationOptions` (`set_options`).
`Edit::prepare()` runs the former `plan()` (bounded source scan, page
generation and readback audit, common insertion plan) and returns a `Plan`
with the `SourceProof` and the plan facts (offsets, lengths, SHA-256s);
`Plan::commit()` (or `Edit::commit()`) returns a `SourceBackedPresentation`
holding the plan, written by `write_to` exactly as the former
`PublicationPlan::write_to` wrote it. Reading or editing a committed
presentation materializes its output once (the 0356a deferred shape);
`Diagnostics { changed(), proof() }`. The `Patch` is exact-pair: it replays
onto a share of the presentation it was committed on, its inverse onto a
share of the commit's presentation (restoring the source bytes exactly), and
anything else is `tail_append::Error::PatchConflict`. `SourceBackedPresentation`
retains its `content.xml` and `styles.xml` as shared `Arc<str>` so committed
presentations share them.

**Gate.** The four families are registered (`litchi_ods::source_backed::{cells,
dde, sheet_metadata}`, `litchi_odp::tail_append`), `SourceBackedSpreadsheet`
and `SourceBackedPresentation` are immutable owners, and the former names
are banned (`SourceCell*`, `SourceMetadata*`, `dde::Source{Edit, Commit,
Patch, Snapshot}`, `apply_dde_patch`, `apply_sheet_metadata_patch`,
`begin_cell_edit`, `cell_snapshot`, `write_to_default`, the `Publication`
suffix under `source_backed`, `tail_append::{PublicationPlan, Edit::plan,
Edit::from_read_at, Edit::with_limits}`). The W6 line of
`docs/API_GLOSSARY.md` §2a is green.

**Consumers.** `tools/perf-baseline` (the ODS source-cell selectors, the ODP
source-tail-append lifecycle and `odp_native_tail_append_probe`) and every
test. The ODP lifecycle now opens a `SourceBackedPresentation` (which
validates and retains `content.xml`) instead of a bare
`SourceBackedPackage`; its timing scope says so.

## 3. Deleted capabilities and their replacement routes (item 2)

| Deleted (0358 or 0358a) | Replacement route | Proof |
|---|---|---|
| ODT `header_footer::{Transaction, Commit, Patch}` | `Package::edit()` layout verbs (`add_master_page`, `set_header_footer_*`, `insert/replace/remove_section_layout`, page-layout verbs); package `Patch`, `inverse`, `Package::apply` | `tests/layout_family_routes.rs` |
| ODT `header_footer::DurablePatch` | layout verbs are durable: a patch with a document step serializes as the exact target artifact (`RestoreSnapshot`), `Patch::durable()` → `transaction::DurablePatch::{to,from}_deterministic_json`, `apply`, `inverse` | `a_layout_edit_has_a_durable_reversible_patch` |
| ODT `header_footer::History` | `Package::history(HistoryLimits)` records layout edits (undo/redo) | `layout_edits_record_in_bounded_history` |
| ODT `header_footer::{MergePlan, Resolution, Target}` | `transaction::MergePlan` over `JoinedEdits`; layout verbs write one owner effect each (`/styles/master-page/<name>`, `/styles/page-layout/<name>`, `/styles/section/<name>`), so disjoint owners merge and one owner conflicts, resolved by `MergeChoice` | `disjoint_layout_owners_merge_and_one_owner_conflicts` |
| ODT layout transfer between documents | `Document::layouts()` → `Snapshot::prepare_{master_page,section_layout}_transfer` → `Edit::insert_{master_page,section_layout}_transfer` | `layouts_and_protection_transfer_to_another_document` |
| ODT `protection` policy `Transaction`, its durable patch, transfer and merge | `Edit::set_protection` (durable), `Document::protection()` + `set_protection` on the destination (transfer), `MergePlan` conflicts on the protection owner | same test, `conflicting_protection_branches_need_an_explicit_choice` |
| ODT `MutableDocument` verbs | 121 `Package::edit()` verbs (0358 §3); `Builder` for fresh documents | `tests/internal/document_verbs.rs` |
| ODS part `Commit`/`Patch` (content validation, database ranges, scenarios, table templates, DDE, sheet metadata, settings, metadata) and `apply_content_validation_patch` | `document::Edit` closures and setters; `document::Patch::{to,from}_deterministic_json` (durable), `inverse`, `join`, `three_way` | `tests/part_family_routes.rs` |
| ODS `MutableSpreadsheet` | `builder::Workbook` | 0358 §4 |
| ODS tracked-change/content-validation `*_with_limits` twins | the public forms with default limits; the parsers take required limits | 0358 §4 |
| ODS `SourceCell{Snapshot, Edit, Commit, Patch, PublicationReport}`, `begin_cell_edit`, `cell_snapshot` | `source_backed::cells`; reads on the spreadsheet itself; `changed_cells` on `cells::Diagnostics`, the byte count and no-op flag on the `write_to` report | `tests/source_cell_transactions.rs`, `tests/source_hyperlink_transactions.rs` |
| ODS `dde::Source{Snapshot, Edit, Commit, Patch}`, `apply_dde_patch`, the DDE model patch | `source_backed::dde`, `apply_dde`, `apply_dde_with` (destination policy) | `tests/internal/ods_dde_facades.rs` |
| ODS `sheet_metadata::SourceMetadata{Snapshot, Edit, Commit, Patch}`, `apply_sheet_metadata_patch` | `source_backed::sheet_metadata`, `apply_sheet_metadata`, `apply_sheet_metadata_with` | `tests/internal/ods_sheet_metadata_transactions.rs`, `tests/internal/sheet_metadata_context.rs` |
| ODP chart `Edit`/`Commit`/`Patch`/`Diagnostics`, inventory `Transaction`/`Editor` | `Package::edit()` chart verbs; `edit::Patch::{to_durable_bytes, from_durable_bytes}`, `inverse`, `join`, `three_way` | `charts/tests.rs::package_chart_edits_have_durable_patches` |
| ODP `tail_append::{Edit::from_read_at, Edit::new, Edit::with_limits, Edit::plan, PublicationPlan}` | `SourceBackedPresentation::edit_tail_append`, `Edit::{set_limits, set_options, prepare, commit}`, `Plan` | `source_append.rs` unit tests incl. `commit_is_deferred_reversible_and_chains` |

Behavior differences, all on purpose:

- a family `Edit::commit` consumes the edit (the 0351 shape); the former DDE
  and sheet-metadata `commit(&mut self, &context)` kept a failed edit for a
  retry. The context is fixed when the edit starts (`edit_*_with`).
- a replayed cells patch reports `changed_cells() == 0` (the count belongs
  to the commit that staged the cells).
- the ODP tail-append patch replays onto shares of its own presentations
  only; the former route had no patch at all.

## 4. Feature matrix and refusals (items 3 and 4)

- `crates/litchi-ops/build.rs` emits `cfg(odf_objects)` when an object
  format (ODG, ODC, ODI, ODB, ODF-formula) is enabled; the 177 former
  `feature = "odf-objects"` gates use it, so the umbrella alone builds only
  its dependency. `tools/check_feature_matrix.py --ops` checks
  `--features odf-objects` alone.
- `litchi_core::Error::Refused(edit::Refusal)` is new (the core error was the
  only one without it); litchi-ops classifies it by kind and
  `litchi_odf_common::Error` maps it both ways.
- In the ODF crates every string-carrying `Error::Unsupported(..)`
  construction (about 240, rewritten mechanically and checked by hand) is
  `Error::Refused(Refusal::new(kind, "<crate>.<module>", "<static reason>"))`
  with names, values and paths dropped from the reason; signed or protected
  sources are `RefusalKind::Protected` (litchi-ops `document_protected`),
  refused lossless content `UnsupportedContent`, the rest
  `UnsupportedOperation`. Allocation failures, cancellations and
  execution-policy rejections that were filed as unsupported are
  `Error::Other`. `odf_common::Error::Signed` now converts to a core
  `Refused(Protected)` instead of a string. No construction of
  `Unsupported(String)` remains in the ODF crates; match arms accept
  `Unsupported(_) | Refused(_)`. The 0289 fidelity report was re-blessed
  for the new texts (codes unchanged except the signed ODT snapshot,
  `document_protected`). An ODT text-set refusal names the replaced ODF
  element again (`text:date`, `text:ruby`, …) through static per-element
  reasons, so the diagnostic is kept without carrying document content.

## 5. Commits

| Commit | Scope |
|---|---|
| `81e4e8705c` | merge trunk `7f36f7b909` (0356a, 0357a) |
| `0a86513faf` | ops: `--features odf-objects` builds on its own |
| `e4a0614f5d` | content-free refusals; durable/history/merge layout edits; package routes for the removed part patches; their tests |
| `77ab7cab8a` | source-backed families (ODS, ODP), the owners' `write_to`/`save`, the ODT element-named refusals, gate, consumers, `odf_source_backed_probe` |
| `a04f5683cf` | this record (draft), A-B-B-A evidence, `p0358a` bench-matrix rows |
| `129c3861a4` | merge trunk `ff0f1fa036` (0358 merged, 0356b, 0357b) |
| this commit | record, rustfmt of `kept.rs`, the side-workspace lock files (trunk's 0357b `sha2` dependency of litchi-ppt) |

## 6. Performance

Two probes, release, pinned to one CPU, run A-B-B-A
(`results/0358a/abba.sh`, `analyze.py`, `summary.json`, `log.txt`; binary
hashes in `binaries.sha256`). Ratio = median of the after runs / median of
the before runs.

**Changed routes** (`crates/litchi-ops/examples/odf_source_backed_probe.rs`
after, its twin `results/0358a/probe_before.rs` on the base `e4a0614f5d`
before; one source differing only in `mod route`, 15 rounds): a
`SourceBackedSpreadsheet` over a 4.3 MiB ODS (12,000 × 8 string cells,
8.5 MiB `content.xml`, 4 MiB incompressible picture) and a
`SourceBackedPresentation` over a 4.0 MiB ODP (2,000 slides, the same
picture); each lifecycle is open + edit + commit + write to a hashing sink.
Peak heap is the growth of a counting allocator over the lifecycle (the
source bytes are allocated before).

| Route | Before (ms) | After (ms) | After/before | Peak heap before (MiB) | After (MiB) | After/before |
|---|---|---|---|---|---|---|
| ODS 1 cell | 349.3 | 345.6 | 0.989 | 165.2 | 165.2 | 1.000 |
| ODS 100 cells | 1447.7 | 1472.6 | 1.017 | 169.4 | 169.4 | 1.000 |
| ODS DDE source | 240.8 | 248.5 | 1.032 | 75.0 | 75.0 | 1.000 |
| ODS sheet metadata (consolidation) | 405.3 | 407.8 | 1.006 | 264.9 | 238.7 | 0.901 |
| ODP tail append | 13.1 | 13.3 | 1.015 | 5.3 | 5.8 | 1.108 |

Every output is byte-identical between the routes and across all four legs
(the dumped outputs' SHA-256 and every round's length and SipHash agree;
`summary.json` `sb_outputs_sha256_identical`). Time is within the 1.05
budget everywhere. The ODP route holds 0.5 MiB more at its peak: it now
opens a `SourceBackedPresentation`, which validates and retains
`content.xml` (0.56 MiB here), where the former route indexed the archive
only; memory stays bounded by `content.xml` (§7, gap 1).

**The 0358 routes** (`odf_w6_probe`, 0358 head `f755903f45` before, this
head after, 61 rounds):

| Route | Before (µs) | After (µs) | After/before |
|---|---|---|---|
| ODT open + get | 140.2 | 139.9 | 0.998 |
| ODT set | 412.8 | 441.6 | 1.070 |
| ODT batch, 1k sets | 1801.9 | 1844.2 | 1.023 |
| ODS open + get | 1282.2 | 1245.1 | 0.971 |
| ODS set | 11.5 | 11.7 | 1.018 |
| ODS batch, 1k sets | 5146.2 | 5122.4 | 0.995 |
| ODP open + get | 121.3 | 121.1 | 0.999 |
| ODP set | 426.6 | 433.3 | 1.016 |
| ODG open + get | 2312.0 | 2276.7 | 0.985 |
| ODG set | 10156.6 | 10019.5 | 0.987 |
| ODC open + get | 78.3 | 78.1 | 0.997 |
| ODC set | 206.9 | 204.9 | 0.990 |
| flat ODT open + get | 379.7 | 413.9 | 1.090 |
| flat ODS open + get | 2207.0 | 2188.2 | 0.992 |

Two routes miss the budget, ODT set (+7.0 %) and flat ODT open (+9.0 %),
although nothing on their path changed. `results/0358a/codegen.json` holds
the diagnosis: `e4a0614f5d` (items 2–4) is level with the 0358 head, and so
is reverting this commit's only ODT change; callgrind counts the same
instructions on both routes (whole probe +0.012 %, `Odt::from_flat`
55.96 M vs 55.94 M, `Odt::set` 89.35 M vs 89.54 M), while cachegrind sees
+8.8 % simulated conditional mispredictions and moved inlining
(`body::xml::Prefixes::resolve` inlined, `drop_in_place<quick_xml::Event>`
out of line) in the shared XML scanning code. The release profile is fat
LTO split into 16 codegen units; the source-backed code added to the binary
moved that partition. Built with `codegen-units = 1`, the same two routes
measure 1.014 and 1.032 (a two-route probe, `results/0358a/odf_w6_mini.rs`).
This is an explained residual of code generation, not added work.

`tools/unified_bench_matrix.py` gains `p0358a`: one after/before row per
route of both probes (the two misses carry this explanation).

## 7. Gaps and follow-ups

1. **ODP tail append retains `content.xml`.** The family starts from a
   `SourceBackedPresentation`, whose open validates and keeps `content.xml`,
   `styles.xml` and `meta.xml`; the former route opened a bare
   `SourceBackedPackage` and scanned `content.xml` as a stream. Memory stays
   bounded by `content.xml` (+0.5 MiB peak on the probe, §6), the time is
   level. A lazily projected `SourceBackedPresentation` would restore the
   index-only open but move its content validation from open to the first
   read, which the ODP tests pin; follow-up if the index-only footprint is
   wanted.
2. **Codegen-partition residual** on ODT set and flat ODT open in the
   default release build (§6): explained residual; no change to their code.
3. A family `Edit::commit` consumes the edit; the former DDE and
   sheet-metadata commits could be retried with the same staged edit.
4. The ODP tail-append patch replays onto shares of its own presentations
   only (exact pair); a replay onto an independently reopened copy of the
   same bytes is refused with `PatchConflict`.
5. Review item 5 (the debt inventory's name heuristic for builder drafts)
   is the PPTX wave's (0356b).

## 8. Gates

All on the merged head `129c3861a4` (plus this commit's formatting and
lock-file updates), one heavy command at a time (`CARGO_BUILD_JOBS=6`,
`RUST_TEST_THREADS=8`). The performance runs of §6 used `77ab7cab8a`.

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 (after formatting `kept.rs`) |
| `cargo clippy -p litchi-{odf-common,odt,odp,odg,odc,odi,odm,oth,odb,odf-formula,odf} -p litchi-ops -p litchi -p litchi-py -p litchi-cli --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ods --all-features --all-targets -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features odt\|ods\|odp\|odg\|odc\|odi\|odm\|oth\|odb\|odf-formula\|odf-objects] -- -D warnings` (none and each) | 0 |
| `cargo test --no-fail-fast --all-features -p` each ODF crate and `litchi-odf` (litchi-ods `--all-targets`), plus `-p litchi-ods --doc` | 0 (4,533 passed) |
| `cargo test -p litchi-ops --all-features` (lib, every integration test in chunks of 20, doc) and `cargo test -p litchi --all-features` | 0 (0 failed) |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p` the ODF crates | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` | 0 (44 profiles, 3,101 banned names, 452 families, 71 immutable owners) |
| `python3 tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py verify` | 0; 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets`; `cd tools/native-resave && cargo check --locked --all-features`; the three ODF fuzz crates `cargo check --offline --all-targets` | 0; 0 (both after the lock-file update); 0 |
| `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references tools.test_check_crate_boundaries tools.test_non_iwork_gate tools.test_check_api_glossary` | 0 (1,296 tests) |
| `python3 tools/officecli_capability_delta.py --check`; `python3 tools/officecli_corpus.py --check`; `python3 tools/unified_bench_matrix.py --check`; `python3 tools/check_feature_matrix.py`; `python3 tools/check_feature_matrix.py --ops` | 0; 0; 0; 0; 0 |

The ODF debt inventory (`tools/api_debt_inventory.py` over the eleven
crates' rustdoc JSON) has 11 rows, none a source-backed family: the nine
`metadata::Commit` aliases of `litchi_core::edit::Commit` it cannot see
through (the gate checks them in `litchi-core`) and the two
`litchi_core::SubEdit` composition aliases of ODB and ODG. The 0289
fidelity report changed in three rows, the restored element names of
§4; no capability, code or golden otherwise changed.
