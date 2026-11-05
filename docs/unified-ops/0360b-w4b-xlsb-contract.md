# 0360b — W4b: XLSB on the record 0351 contract

- Status: implemented on `uo2/w4b-xlsb`, cut from `uo2/w4-xlsx` at
  `ba96140de3`; merged with `uo2/w4-xlsx` at `635b884c7c` (`b45809564b`),
  with trunk `feat/unified-ops-api` (`9560448a9e`) and with the rewritten
  `uo2/w4-xlsx` tip `698fd6c5f0` (`b72c6a760b`; its XLSX files taken as is,
  `c89d0972c3` keeps trunk's ODP perf harness).
- Date: 2026-10-21.
- Input: owner decision [0350](0350-owner-api-debt-decision.md); contract
  [0351](0351-api-contract.md) §2, §4.3, §5 (wave W4, XLSB half); the
  source-backed family shape of [0355a](0355a-w1-docx-gaps.md) §1; the
  XLSB sheet scenarios of [0353a](0353a-sheet-errors-and-sync-engine.md) §4.
- Owned: `crates/litchi-xlsb/**`; the XLSB hunks of `crates/litchi-ops/**`
  (`format/xlsb*`, the XLSB `Classify` arms, `view/sheet.rs`, tests,
  examples), the facade tests and examples, `tools/perf-baseline`
  (`xlsb_crud`, `xlsb_synthetic_fixture`, the `unified-xlsb-*` harnesses and
  their `Cargo.lock` files); the `xlsb-*` profiles of
  `tools/api_glossary.json`; the XLSB text of `docs/API_GLOSSARY.md` (§2,
  §2a line W4).
- `performance_claim`: measured (§6).

## 1. Summary

- **Owner.** `xlsb::Package` is the immutable owner. `Package::into_workbook`
  is deleted; the read-only `xlsb::Workbook` projection is
  `Package::workbook() -> Result<&Workbook>`: every validating open and every
  commit parses it once and caches it in the package (clones share it); a
  package adopted without validation, or published by a family whose change
  cannot affect it (signature, ribbon, task panes, VBA, theme), parses it on
  first use. Outputs take `&self`; `sign`, `sign_with`, `resign`,
  `resign_with` and `unsign` return a new `Package` (the new signature's
  `PackURI` is read through `signatures()`, as in DOCX). No public method of
  `Package` or `Workbook` takes `&mut self`; both are gate owners.
- **Families.** The 50 mutating `Workbook` verbs are gone. Every scoped edit
  is a family `x::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}` with
  `Package::{x, edit_x, edit_x_with_options, apply_x}` and `Edit::reset`;
  29 families are registered with `family_methods` (§2). A commit publishes a
  new `Package` with the patch and diagnostics; no `Transaction`,
  `rollback`, `is_changed`, snapshot publication or `*_with_limits` name is
  left, and the gate bans them in both `xlsb` profiles.
- **Owned part engine.** The part families (merged cells, page breaks, auto
  filter, hyperlinks, data validations, conditional formats, comments,
  drawings, threaded comments, sparklines, cell watches, slicers, timelines,
  scenarios, task panes, ribbon, VBA, connections, shared workbook, raw OPC)
  stage on one copy-on-write engine (`part_edit::Staged`, ported from XLSX)
  and commit an exact, invertible part patch.
- **Byte codecs.** The detached byte-level APIs keep every capability as
  drafts without a commit type: `cell_values::worksheet::Draft::finish() ->
  (Snapshot, Patch)`, `external_link::Draft`, `web_extension_bindings::Draft`.
- **Errors (§2.4).** `PackageError` gains `Refused(Refusal)` and
  `InvalidLimit { kind, value, maximum }`; `LimitExceeded { resource: &str }`
  is deleted and every bound is `Limit { kind, observed, maximum }` with a
  typed crate `LimitKind` wrapping the owning family's kind. Signed refusals
  are `Error::Signed`. litchi-ops classifies `Refused` by kind,
  `PatchConflict` as `invalid_input` and `InvalidLimit` as
  `invalid_argument`.
- **Limits (§2.5).** `SemanticLimits` (new aggregate of the ten family
  budgets) rides on `OpenOptions::semantic_limits`, the package
  (`semantic_limits()`, `with_semantic_limits()`) and the projection; family
  edits take `x::EditOptions { limits, .. }`; getters are `*_value()`. Raw
  constructors that took limits are `bounded`.
- **Drafts (§2.1 rule 7).** `builder::{Builder, Worksheet, ChartSheet, ..}`
  replaces the public `writer` module and its `Mutable*` names.
- **Source-backed owner.** `SourceBackedWorkbook::edit_metadata` follows the
  0355a §1 shape (an owning `Edit` with `reset`, a commit that validates and
  defers output).

Fidelity: every round-trip, CRUD, corpus and no-op test passes with
byte-identical outputs; formula, calculation and sheet-trait tests pass
unchanged; the OfficeCLI corpus and the capability and registry goldens are
unchanged (no re-bless).

## 2. Families

| Module | Entry points | Engine |
|---|---|---|
| `cell_values` | `edit_cell_values`, `apply_cell_values` | record (existing) |
| `cell_values::structure` | `edit`, `edit_with_options`, `apply` | record (existing); `History::{undo, redo}` take `&Package` |
| `metadata` | `metadata`, `edit_metadata`, `apply_metadata` | record |
| `page_breaks::edit`, `auto_filter::edit`, `hyperlinks::edit`, `data_validation::edit`, `conditional_formatting::edit`, `comments::edit`, `drawing::edit` | `edit_x`, `apply_x` | part (existing families moved off `Workbook`) |
| `opc` | `edit_opc`, `apply_opc` | raw OPC copy; validates the candidate, reparses |
| `merged_cells::edit` | `edit_merged_cells`, `apply_merged_cells` | part |
| `scenarios` | `edit_scenarios`, `apply_scenarios` | part |
| `slicer`, `timeline` | `edit_slicers`, `edit_timelines`, `apply_*` | part; closure validated and read back on every change |
| `task_panes`, `ribbon`, `vba` | `task_panes`/`ribbon`/`vba`, `edit_*`, `apply_*` | part, unparsed publication |
| `package::connections` | `connections`, `edit_connections`, `apply_connections` | part |
| `custom_data` | `custom_data`, `edit_custom_data[_with_options]`, `apply_custom_data` | record |
| `theme` | `theme`, `edit_theme[_with_options]`, `apply_theme` | record, unparsed publication |
| `sparkline::edit`, `cell_watches` | `edit_*`, `apply_*` | part |
| `calculation_chain`, `data_model`, `xml_maps`, `volatile_dependencies` | `x`, `edit_x[_with_options]`, `apply_x` | record |
| `comments::threaded::edit` | `threaded_comments`, `edit_threaded_comments`, `apply_threaded_comments` | part |
| `shared_workbook` | `shared_workbook`, `edit_shared_workbook`, `apply_shared_workbook` | part |

## 3. Decisions

1. **Signature verbs return a `Package`** (the DOCX precedent); the new
   signature part is reported by `signatures()`.
2. **Part-patch replay checks the touched parts.** A part patch carries the
   exact before and after image of every part it changes and of the root
   relationships; replay refuses a package whose touched parts differ, so a
   no-op patch applies to any package and an A-B-A package accepts a patch
   made on A (tests adjusted). Connections, threaded comments and the shared
   workbook now have exact part patches instead of semantic ones.
3. **Lexical-only raw OPC edits are no-ops.** `edit_opc` compares parts by
   content type, bytes and relationships; a change that only respells a
   relationship part (same relationships) publishes nothing. The theme
   stale-relationship test opens the respelled graph as a new package.
4. **Timelines over a non-canonical source** refuse every non-empty change
   with `Refused(UnsupportedContent)`; removal stays allowed.
5. **Readback.** The connections commit compares the stored catalog with the
   canonical form `store` returns; the shared-workbook commit checks that the
   result reads back (it no longer compares with the staged catalog, which
   the store normalizes), as before the wave.
6. **Raw `new` constructors with default limits stay** (the PPT §2.5.4
   precedent); the limit-taking ones are `bounded`.
7. **Hard-ceiling validation is `InvalidLimit`.** A configured calculation
   chain or volatile dependencies budget above its implementation ceiling is
   an invalid limit, not a limit refusal.
8. **litchi-ops eager owner.** `XlsbOwner::Eager` holds the `Package`; it is
   admitted only after its projection is parsed, so reads borrow the cached
   projection, edits hand the package itself to the families (no
   `Workbook::to_package` copy) and saves write the package directly.

## 4. Gaps

1. **Free-text refusals.** About 250 `UnsupportedFeature(String)` and
   `Unrecognized` sites, mostly on read paths, are not yet
   `Refused(RefusalKind)`; only the timeline refusal moved. Follow-up: a
   0360ba pass over the edit-path sites (the read-path ones keep the
   corrupt-file class).
2. **Content in messages.** `WorksheetNotFound(name)` and a few
   `InvalidFormat` messages still carry user content (ADR 0034 §6 audit,
   item 22, stays deferred as in 0351).
3. **Source-backed metadata patch replay.** `SourceBackedWorkbook` has no
   `apply_metadata`; a metadata patch replays on an owning `Package`.
4. **ooxml-common and XLDM limits.** Their limit errors still carry a
   resource string; XLSB maps them to `CustomData::PropertiesXmlBytes` and
   `DataModel::Model`.
5. **Lexical-only OPC edits** (decision 3) are dropped by `edit_opc`.

## 5. Commits

| Commit | Scope |
|---|---|
| `5ecc2fa900` | immutable `Package`, the families, `SemanticLimits`, `builder`, error variants, consumers |
| `bc370606fd` | typed limit kinds, gate registration, glossary statuses |
| `b45809564b`, `9560448a9e` | merges of `uo2/w4-xlsx` and trunk |
| `be1538b639` | owning source-backed metadata `Edit` with `reset`; `xlsb_w4b_probe` |
| `1af8b42cb6` | shared package graph; part patches capture manifests on first replay |
| `4e0d748bf7` | record, evidence, `p0360b` |
| `b72c6a760b`, `c89d0972c3` | merge of the rewritten `uo2/w4-xlsx` tip; trunk's ODP harness kept |

## 6. Performance

`crates/litchi-ops/examples/xlsb_w4b_probe.rs` times the §5 W4 XLSB routes:
leaf legs with only the XLSB calls the wave keeps
(`Package::{new, from_bytes, edit_cell_values, to_bytes}`,
`cell_values::Edit::{set_scalar, commit}`) and unified legs with only the
`litchi_ops` API, so one source builds on both trees. The 0353a XLSB sheet
scenarios run `sheet_dispatch_bench` (unchanged source), and the leaf CRUD
route runs perf-baseline `xlsb_crud --case all --backend owned` (each
tree's own source; the base copy needs the one-line `Cell` trait import
this wave also adds, without which it did not build). Base: `git archive`
of the merge tree `684a028b2a` (`uo2/w4-xlsx` `635b884c7c` + trunk, i.e.
the head without this wave's commits); head: `1af8b42cb6`. Release (fat
LTO), pinned to CPU 29, 31 rounds, A-B-B-A twice
(`results/0360b/{abba.sh, build.sh, analyze.py}`; run 1 in
`results/0360b/`, run 2 in `results/0360b/run2/`; binary hashes in
`binaries.sha256`, binaries not committed). Ratio = median of the after
runs / median of the before runs (run 1; run 2 in parentheses).

| Route | Before (µs) | After (µs) | After/before |
|---|---:|---:|---:|
| leaf `Package::from_bytes` (100×10) | 72.4 | 70.8 | 0.978 (0.978) |
| leaf one-scalar commit | 424.2 | 432.9 | 1.020 (1.037) |
| leaf one-scalar commit + `to_bytes` | 423.6 | 427.0 | 1.008 (1.043) |
| leaf 1k-scalar commit | 1316.6 | 1255.2 | 0.953 (0.969) |
| `File` open + get (source / eager) | 159.0 / 191.7 | 164.0 / 183.4 | 1.032 (1.011) / 0.957 (0.942) |
| `File::set` value (source / eager) | 762.6 / 413.7 | 705.9 / 353.3 | 0.926 (0.926) / 0.854 (0.852) |
| `File::set` + `to_bytes` (source / eager) | 760.0 / 411.3 | 706.7 / 351.1 | 0.930 (0.927) / 0.853 (0.845) |
| `File::set` bold (source / eager) | 1649.6 / 1167.0 | 1570.4 / 1075.4 | 0.952 (0.926) / 0.921 (0.894) |
| `/metadata` title (source / eager) | 18.6 / 35.0 | 18.7 / 35.4 | 1.001 (1.006) / 1.011 (1.000) |
| `File::edit`, 1k sets (source / eager) | 2554.7 / 1987.0 | 2465.7 / 1897.7 | 0.965 (0.957) / 0.955 (0.945) |
| add note / hyperlink / name (eager) | 45.9 / 40.6 / 1765.9 | 47.5 / 41.9 / 1829.5 | 1.034 (1.035) / 1.033 (1.034) / 1.036 (0.968) |
| merge / two chained sets (eager) | 963.6 / 2960.0 | 957.5 / 2774.3 | 0.994 (0.925) / 0.937 (0.948) |
| 0353a `cell` / `cell_value` / `cells` | 799.6 / 799.3 / 296.9 | 828.3 / 802.2 / 293.5 | 1.036 (0.996) / 1.004 (0.991) / 0.989 (0.991) |
| 0353a `evaluate` / `recalculate` / `recalculate_heavy` | 1739.1 / 5355.7 / 6368.3 | 1761.2 / 5447.6 / 5970.4 | 1.013 (0.987) / 1.017 (1.047) / 0.938 (0.946) |
| `xlsb_crud` open, catalog, cell, scan, text | 975.5 / 517.9 / 526.4 / 522.5 / 164.5 | 643.4 / 513.9 / 520.0 / 519.7 / 163.0 | 0.660 / 0.992 / 0.988 / 0.995 / 0.990 |
| `xlsb_crud` no-op commit save, one edit save, 1% edit save | 924.0 / 2486.4 / 2485.6 | 573.9 / 2145.1 / 2148.0 | 0.621 / 0.863 / 0.864 |

Every route is within the 1.05 budget in both runs (largest 1.047,
`recalculate` run 2, 1.017 in run 1; the wave does not touch evaluation).
The batch budget: the head's `File::edit` of 1,000 cell sets is 1.51
(1.52) × the one-commit leaf edit on the eager owner and 1.96 (1.97) × on
the source-backed owner (base 1.51/1.94), within 2.0. The eager `set` is
15 % faster because the eager owner now hands its package to the family
instead of copying a `Workbook` into a new package; the `xlsb_crud` open
and no-op save gains come from the shared, cached projection (the base
reparsed the workbook in `into_workbook`); neither is otherwise claimed.

Profiling (callgrind, `LITCHI_PROBE_ONLY`): before the last perf commit the
note add cost 1.1 M instructions against 0.38 M on the base; the copied
XLSX engine serialized and validated both sides' content-types and
relationship manifests on every commit. `1af8b42cb6` shares the package
graph (`Arc<OpcPackage>`), keeps the two graphs in the patch and captures
the manifests on first replay only, and only the views a patch can change;
the note add is 0.36 M after it.

`tools/unified_bench_matrix.py` gains `p0360b` (37 rows, every one meets).

## 7. Gates

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-xlsb -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 (after the `698fd6c5f0` merge; before it, 101 on one `needless_borrow` in the merged `litchi-xlsx` form-control code, with `--no-deps` 0) |
| `cargo clippy --no-deps -p litchi-ops --no-default-features --all-targets -- -D warnings`; the same `--features xlsb` | 0; 0 |
| `cargo clippy -p litchi-xlsb --all-targets` with no features, `sign`, `vba-inspection`, all features | 0 each |
| `cargo test -p litchi-xlsb --all-features` | 0 (56 targets) |
| `cargo test -p litchi-ops --all-features`: `--lib`, `--doc`, the 14 `--test` chunks | 0 each (956 lib, 215 doc tests) |
| `cargo test -p litchi --all-features` | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-xlsb -p litchi-ops -p litchi` | 0 |
| `python3 tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py verify` | 0; 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` (48 profiles; `xlsb-minimal`/`xlsb-full`: 28/29 families, `family_methods`, 2 owners) | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 (301 tests) |
| `cd tools/perf-baseline && cargo check --locked --all-targets`; with `--features xlsb-crud`; the eleven `unified-*` harnesses (`--offline`, locks refreshed) | 0; 0; 0 |
| `cd tools/native-resave && cargo check --locked --all-features` | 0 |
| capability and registry goldens (`--test capabilities --test registry`) | 0 without `LITCHI_BLESS`: unchanged, no re-bless, delta table unchanged |
| `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references tools.test_check_crate_boundaries tools.test_non_iwork_gate` | 0 (1,269 tests) |
| `officecli_capability_delta.py --check`, `officecli_corpus.py --check`, `unified_bench_matrix.py --check`, `check_feature_matrix.py` (31 builds) | 0 each |

The OfficeCLI corpus needs no re-bless: no behaviour visible to it changed.

## 8. Log entry

- W4b XLSB: `Package` is the immutable owner, `Workbook` its read-only projection (`Package::workbook()`); `into_workbook` and the 50 mutating `Workbook` verbs are gone.
- 29 families on the 0351 §2.3 shape registered with `family_methods`; no `Transaction`/`rollback`/`is_changed`/snapshot publication/`*_with_limits`; `builder::` drafts replace `writer::Mutable*`.
- Errors: `Refused`, `InvalidLimit`, typed `Limit { kind }` replacing `LimitExceeded { resource }`; `Signed`; litchi-ops classifies by kind.
- Consumers ported (ops, facade, perf-baseline, unified harnesses); every round-trip, corpus and golden unchanged.
- ABBA twice: all routes ≤ 1.05 (eager set 0.85), 1k batch 1.52/1.97 × leaf; `p0360b`; every gate 0.
