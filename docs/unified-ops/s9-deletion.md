# S9 — delete the format-family types and the migration scaffolding

- Status: done on `uo/m0a-s9-delete`, branched from `uo/m0a` at `0d228d293a`
  (S1–S8d merged).
- Date: 2026-10-05
- Scope: plan [0003](0003-m0a-migration-plan.md) §4 step S9, the S9 items in
  [`m0a-log.md`](m0a-log.md) (S2, S4, S7, S8a, S8c, S8d), and the S8c record's
  notes for S9 ([`s8c-facade-tests.md`](s8c-facade-tests.md) §5).
- Commits: `59bc8f8772` (the deletion) and `27e35692bf` (the close-out
  fixes of §3, items 1 and 2).

After this step, `litchi::open*`, which returns `litchi::File`, is the
facade's only entry point to the formats that `File` covers. iWork keeps
`litchi::iwork`, and the text workbooks keep `sheet::text`. The following
are deleted:

- `litchi::Document`, `litchi::Presentation` and `litchi::sheet::Workbook`,
  with their element types and openers;
- the facade's own Markdown adapter for those types;
- `litchi_ops::__migration`;
- the 47 harness routes that timed the deleted types.

No item in `litchi-ops` or in the facade's `src/` is `#[doc(hidden)]` any
more.

## 1. Deleted

Line counts are `wc -l` of the files at the base commit. Not counting this
record, the branch changes 76 files, with 1,803 lines added and 27,382
deleted. Of the deletions, 19,314 lines are 26 whole files.

### Facade (`crates/litchi`)

| Path | Lines | Contents |
|---|---:|---|
| `src/document/` (8 files) | 6,194 | `Document`, `DocumentElement`, `Paragraph`, `Run`, `Table`, `Row`, `Cell`, `CellMerge` |
| `src/presentation/` (4 files) | 3,397 | `Presentation`, `Slide`, the old `SlideDescriptor` |
| `src/sheet/workbook.rs` | 2,635 | `sheet::Workbook` |
| `src/sheet/workbook_types.rs` | 213 | its owner enum and helpers |
| `src/sheet/adapters/` (3 files) | 1,683 | the XLSX and XLSB `WorkbookTrait` adapters behind it |
| `src/sheet/selection.rs` | 153 | `SelectedWorksheet`, `SelectedCell`, `SelectedCellView`, `WorksheetSelector` |
| `src/sheet/functions.rs` | 477 | 30 public openers: 24 return `Box<dyn WorkbookTrait>`, 6 (`open_xls_workbook*`, `open_xlsb_workbook*`) return the leaf workbooks. Also 3 private helpers. |
| `src/markdown/{document,docx,presentation,writer}.rs` | 3,052 | the old Markdown adapter for the deleted types |
| `tests/markdown_file_parity.rs` | 691 | the parity test that pinned the goldens to the old types |
| **Total (24 files)** | **18,495** | |

Plan 0003 counts "all 42" openers in `functions.rs`. At both the plan commit
`a2e49fac5a` and the base, the file held 30 public functions and 3 private
helpers. The whole file is deleted.

Edits in files that stay:

- `src/lib.rs`
  - Removed:
    - the `document` and `presentation` modules;
    - the root re-exports of `Document`, `DocumentElement`, `Presentation`
      and `Workbook`;
    - the crate doc's "Older entry points" section;
    - `map_ooxml_error`. It was defined here and called on 114 lines in 9
      files, all of them deleted.
  - `litchi::SlideDescriptor` now names
    `litchi_ops::format::pptx::SlideDescriptor`, the descriptor that
    `Pptx::slide_catalog` returns. Its accessors are unchanged (`position()`,
    `slide_id()`), and it now also derives `Hash`.
- `src/sheet/mod.rs` keeps:
  - `view`, `traits`, `types`, `eval`, `FormulaEvaluator` and `text`;
  - `Cell`, `CellIterator`, `RowIterator`, `WorkbookTrait`, `Worksheet`,
    `WorksheetIterator`, `CellValue` and `Result`.

  The coordinate re-exports `sheet::{Address, Area, At, Rect}` are removed.
  They arrived with the deleted selector API (`45c6c625c1`), and the same
  types stay at `litchi::xlsx::*`.
- `src/markdown/mod.rs` keeps only the `litchi-markdown` re-exports. The
  conversion for `File` lives in `litchi-ops`, as before.
- `Cargo.toml` drops the optional dependencies `base64`, `chrono` and
  `memchr`, which only the deleted code used:
  - `rtf` no longer enables `chrono`;
  - `markdown` no longer enables `base64` or `memchr` (the ops `markdown`
    feature brings its own).
- `src/detection_smart.rs` calls the public `litchi_ops::detect::*_with_tail`
  functions. The dead `detected::detect_ooxml_path_with_limits` is removed.
  The probe-counter import is gated to the OOXML test builds that read it,
  which removes a base warning about an unused import.
- Documentation:
  - `README.md` (the crate README) gets its Usage and Entry Points rewritten
    for `litchi::open` and `File`. S8d had rewritten the root README and the
    guide, but not this file.
  - `src/tolerance.rs`, `src/iwork/mod.rs` and the headers of
    `tests/markdown_corpus_golden.rs` and `tests/unified_open_facade.rs` drop
    their "until S9" wording. The goldens are fixed reference data now: a
    difference is a behaviour change to review, not output to bless.

### `litchi-ops`

- `src/__migration.rs` (24 lines) and its `#[doc(hidden)] pub mod` are
  deleted.
- `src/detect/detected.rs` shrinks from 4,261 to 882 lines. It held the
  per-family source detectors that only the deleted types called, through
  `__migration`. Once `__migration` was gone, the compiler flagged many of
  them as dead. Others carried `#[allow(dead_code)]`, so reachability was
  traced by hand from the live roots: the `open*` functions and the public
  `detect` API. Deleted:
  - 23 public functions: the `detect_{document,presentation,workbook}_source_*`,
    `detect_{docx,pptx}_source_*`, `detect_prepared_{odp,ods}*`,
    `detect_odt_source_path_with_limits`, `detect_ooxml_path_with_limits`,
    `read_*_path_bytes_with_limits`, `try_open_xls_source` and
    `classify_ole_host_stream` families;
  - 11 public result enums;
  - 3 `UNIFIED_*_FALLBACK_MAX_INPUT_BYTES` constants;
  - `PreparedOdfProbeError`;
  - 29 private helpers.

  `detect_and_prepare_path_with_limits` and `OdtSourcePathCandidate` serve the
  live open route and are now `pub(crate)`.
- Visibility:
  - `detect/workbook_metadata.rs` (`MetadataRepairs`,
    `read_source_backed_metadata`) and `view::PresentationView::extract_text_for_markdown`
    become `pub(crate)`. The latter was `#[doc(hidden)] pub`, and its cfg now
    names the builds that use it.
  - `lib.rs` loses the private leaf aliases `docx`, `ppt`, `xls`, `xlsb` and
    `xlsx`, which only the deleted detectors used.

### Harness (`tools/perf-baseline`)

The 47 predecessor selectors that timed the deleted types are removed. They
are listed by family in the harness README's new subsection "Predecessors
removed in M0a step S9". The registry goes from 652 to 605 names.

| File | Change |
|---|---|
| `src/facade_ole2.rs` | deleted (795 lines). Change 0638's DOC/PPT facade family. Its corpus generator identity moves to `ole2_range_source.rs` (`DOC_REAL_FILE_GENERATOR`, same string), which the S8b OLE2 pairs still record. |
| `src/lib.rs` (72,232 → 70,184 lines) | Removed from the `Case` enum, `Case::name`, `parse_case`, `--help`, the predicates and the dispatch:<br>- the 47 variants;<br>- the RTF, ODT, ODP and XLSX-bytes root runners and their helpers;<br>- the `SourceSummary` fields `facade_ole2`, `rtf_root`, `odt_root` and `odp_root`, which only the deleted selectors filled, with their types. Three of those types (`RtfRootSummary`, `OdtFileRootSummary`, `OdpRootSourceSummary`) derive `Default`, which hid them from the dead-code lint.<br>`EXPECTED_SELECTABLE_CASE_COUNT` = 605. A new test, `m0a_s9_deleted_the_facade_predecessor_selectors`, checks that all 47 names stay unparseable. |
| `src/filesystem.rs` (8,825 → 7,844 lines) | `Operation` loses the predecessor variants and their runners and verifiers. The S8b source-backed routes still replay their predecessor's untimed leaf pass, now keyed by a new `SourceReplay` descriptor, whose evidence still names the predecessor. Six tests are ported to the unified routes. |
| `src/filesystem/ops.rs`, `src/unified_ops.rs` | `Route::source_replay()` returns `SourceReplay`. The route catalog names each predecessor by its frozen selector string, because the `Case` value is gone. |
| ODS root runner (in `lib.rs`) | `ods_file_{eager,source}_open` are removed. The four typed routes `ods_file_{eager,source}_selected_{cell,media}` stay: they time only the ODS leaf owners, and `File` has no ODS cell or media read in M0a. Their untimed workbook expectations now come from `litchi_ods::Spreadsheet`, which reports the same values the deleted facade workbook did. |
| `src/bin/xlsb_crud.rs` | `open_identify` times the leaf call that the deleted `sheet::open_xlsb_workbook_from_bytes` wrapped. `full_text` times `litchi::open_bytes` + `File::text`, the unified successor of `sheet::Workbook::from_bytes(..).text()` (see §6). |
| `src/bin/xls_source_attribution.rs` | The `facade-file` mode is removed. The other modes' `comparison_scope` strings still mention it, so that the output stays the same. |
| `src/bin/docx_source_selected_paragraph.rs` | The untimed facade oracle reads `litchi::open` + `File::document`. Its stale-source check accepts the `SourceChanged` leaf of whichever layer noticed the change. |
| `tests/xlsx_filesystem.rs` | The fresh-child and cold-verified checks run on the unified XLSX successors, which report the same evidence. |
| `README.md` | The new S9 subsection, with the removed and kept selectors and their successors. 14 one-line markers at the passages that describe removed selectors; the passages themselves stay, because they describe the frozen records. |
| `Cargo.lock` | The `litchi` entry loses `chrono`. The lock was refreshed offline; nothing else changed. |

`docs/performance/crud-coverage-index-v2.json` binds the 605 names in
`Case::name` order:

- digest `c0e60833f573fdf289c1444a0ddc8cd1b8733d8f6791ad7a29fb248cd9f1a76b`;
- 572 names excluded, 33 mapped.

All 47 removed names were in the `excluded` list. `tools/test_crud_coverage_index.py`
pins 605 and the digest. In `tools/test_perf_baseline_source_policy.py`:

- the two `facade_ole2.rs` tests become an ordinary-save hygiene test and a
  test that `--ole2-file` is the single input for DOC and PPT;
- the paired-route test now requires each predecessor to be either one of the
  47 deleted names (and unregistered) or one of the 15 kept ones.

Frozen records under `docs/performance/` are untouched and keep the deleted
names.

## 2. Kept, and why

| Item | Why |
|---|---|
| `litchi::common::Error` | It is `litchi_core::Error`, re-exported by `pub use litchi_core::*`. It is the typed leaf `Source::Core` of the unified error, the error type of `images` and `MetadataYaml`, and tests name it. |
| `litchi::sheet::text` | The concrete `TextWorkbook` and `formats::read_*`, as the plan requires. |
| `litchi::iwork` and the leaf re-export modules | as the plan requires |
| The detection, `sheet` and `markdown` re-exports the tests use | `common::detection::*`, `detection_smart::{ooxml, …}`, `sheet::{CellValue, FormulaEvaluator, WorkbookTrait}`, `markdown::{MarkdownOptions, TableStyle, ToMarkdown}` (S8c §5) |
| `litchi_ops::detect::tail::{ZipTail, NoZipTail}` | Now a public, documented module. The facade's iWork classifier implements `ZipTail`. |
| `litchi_ops::detect::*_with_tail`, `litchi_ops::{open_with_tail, open_bytes_with_tail}` | Public and documented, with their `# Errors` and when the tail runs. Each was `#[doc(hidden)]` or reached only through `__migration`. |
| `litchi_ops::detect::{opc_probe_count, reset_opc_probe_count}` | The probe counter's test path, under the `test-support` feature. It is no longer `#[doc(hidden)]`, and a doctest shows one OPC index per DOCX open. |
| 15 predecessor selectors | `detect_{odt,ods,odp}_{bytes,reader,polyglot,catalog_alias}` (12), `docx_semantic_one_paragraph`, `docx_semantic_one_paragraph_text` and `docx_producer_source_selected_paragraph`. They time the detection functions or the DOCX leaf owner, not the deleted types. |
| `ods_file_{eager,source}_selected_{cell,media}` | They have no unified successor in M0a (S8b), and their timed scope never touched the deleted workbook. |
| The DOCX and PPTX source replays in `filesystem.rs` | The S8b paired source-backed routes replay their predecessor's untimed leaf pass and record it under the predecessor's name. They are keyed by `SourceReplay` now, not by the deleted `Operation` variants. |

## 3. Close-out fixes

1. **XLSX coordinate and range errors** (S8d finding). The XLSX leaf reports
   a caller's bad coordinate with the same `Coordinate` and `Range` errors it
   uses for a bad reference inside a worksheet. So the classifier cannot tell
   the two apart, and both fell through to `corrupt_file`. The input is now
   resolved where ops receives it, and the code is set there. The typed leaf
   and the message are kept, and the suggestion is cleared. The new helper is
   `Error::caused_by_argument`, which is `pub(crate)`.
   - `SelectedWorksheet::cell`:
     - an A1 reference that names no grid cell → `invalid_argument`;
     - indices outside the grid → `out_of_range`.
   - `SelectedWorksheet::cells`:
     - bounds inside the grid that enclose no cells, or a bad A1 range →
       `invalid_argument`;
     - bounds that leave the grid → `out_of_range`.
   - The `WorkbookTrait` worksheet (`cell`, `cell_value`, `row`,
     `cell_by_coordinate`) → `out_of_range` or `invalid_argument` in the same
     way. A row index above `u32::MAX` is `out_of_range`.
   - A coordinate read from a malformed worksheet stays `corrupt_file`.
     `sheet_trait_coordinates_are_argument_errors_and_payload_coordinates_corruption`
     pins both sides, and the selector tests in ops and in the facade's
     `unified_sheet_selectors.rs` assert the codes. The error module's
     classification table names both cases.

   XLSB coordinates are unchanged (§6).
2. **CryptoAPI `.ppt`** (S8a finding). Without its `encryption` feature, the
   PowerPoint reader cannot decrypt, so it reported RC4 CryptoAPI (PowerPoint's
   only password-to-open scheme) as `UnsupportedEncryption(CryptoApi)` instead
   of `PasswordRequired`. The file therefore opened as `unsupported_capability`
   rather than as `File::Locked`.
   - The fix is `format::ppt::password_required`, used by both the byte route
     and the path route. Without the ops `encryption` feature, it treats that
     refusal as "encrypted". With the feature, the reader asks for the
     password itself, and an unsupported scheme stays unsupported.
   - Evidence: `tests::security::locked::encrypted_legacy_documents_open_locked_with_their_scheme`
     failed in a `ppt`-only build at the base and passes now (`cargo test -p
     litchi-ops --no-default-features --features ppt --lib`, 64 passed).
   - The Python oracle, which builds without `encryption`, adds
     `Password_Protected-hello.ppt` to its `ProtectedError` cases.
3. **S4 doc comments** (S4 finding). These comments said that ops-only
   detectors classify iWork:
   - the `detect` module docs;
   - `detect::tail`;
   - `detect_and_prepare` ("RTF and iWork are classified");
   - the "Finally try iWork detection" comment;
   - `DetectedFormat`'s variant list.

   They now say that a detector without a tail reports an iWork package as
   unrecognized, and that the `_with_tail` forms take the classifier. The
   detectors consult a tail only in builds with an iWork feature. The openers
   consult it in every build. The `tail` docs state both.
4. **`Display` is facade-level** (S2 follow-up). No change was needed. Since S6
   (`eca894f849`), `Display` writes the code, its summary, the format and the
   suggestion, never the leaf text. `error/tests.rs::display_does_not_repeat_the_leaf_text`
   pins this. The leaf is reachable through `Error::leaf()` and `source()`.
5. **Rustdoc links.**
   - `litchi::iwork`'s `[`Document`]` resolved to the root `Document` that
     this step deletes. It now names `crate::iwork::Document` explicitly.
   - The crate doc's `[`markdown::ToMarkdown`]` link (from S7) warned in every
     build without `markdown`, for example in the rustdoc build of
     `tools/check_iwork_public_api.py`. It is a code span now.

## 4. Tests

| Crate | `#[test]` before → after | Where the difference is |
|---|---|---|
| `litchi` | 549 → 233 | 304 in the deleted `src/` modules (S6 had ported their subjects to `File` tests, S8c §5.3), and 12 in `markdown_file_parity.rs` |
| `litchi-ops` | 401 → 390 | `detect/detected.rs` 27 → 5; `tests/open.rs` +10; `format/xlsx/tests.rs` +1 |
| `tools/perf-baseline` | 671 → 659 | `facade_ole2.rs` −5; `lib.rs` 180 → 174 (tests of deleted selectors, plus the new guard); `xls_source_attribution` 11 → 10 |

The 22 removed detector tests:

- 17 are ported onto `crate::open*` in a new `tests/open.rs` module,
  `zip_candidates` (10 tests). The ported tests cover:
  - the ODS and ODP prepared-package handoff;
  - OOXML precedence for an uncertain alias catalog;
  - terminal classification of a valid non-DOCX OPC package;
  - the DOCX owner's refusal;
  - ZIP inputs without an Office main part;
  - the exact OOXML read limit;
  - a malformed ZIP candidate.
- 5 are dropped with the functions they tested: the two unit tests of
  `effective_presentation_path_input_limit`, and the three tests that the
  per-family byte probes handed the same allocation back for fallback.

The per-format test helpers (DOCX, ODP, ODS, ODT, PPT, PPTX, XLS, XLSB, XLSX)
now open through `crate::open*` instead of the deleted detectors.

## 5. Gates

Every cargo command ran with `CARGO_TARGET_DIR=<worktree>/target
CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0`,
from the worktree root. The worktree is sparse. The Python suites that read
`docs/performance/results/` ran with exactly the files they read added to the
sparse checkout, as in S8b, and the original patterns were restored
afterwards.

The table's gates ran on the tree that `27e35692bf` commits. The
deletion-only commit `59bc8f8772` was checked on its own as well:

- clippy `-D warnings` with all features on `litchi` and `litchi-ops`;
- `cargo check -p litchi-ops --all-targets` with `ppt`, `xlsx` and `docx`;
- the ops unit tests (387);
- `unified_sheet_selectors`;
- `cargo check` of the harness.

| Gate | Command | Result |
|---|---|---|
| Format (workspace) | `cargo fmt --all --check` | Exit 1, with the same 11 hunks as the base, all in files this step does not touch: `litchi-iwa-archive/src/package.rs` (8), `litchi-iwa-archive/src/zip.rs` (2), `soapberry-zip/src/office.rs` (1) |
| Format (harness) | `cargo fmt --check` in `tools/perf-baseline` | clean |
| Clippy | `cargo clippy -p <p> --all-targets --all-features -- -D warnings` for `litchi`, `litchi-ops`, `litchi-py` | clean, 3 of 3 |
| Feature matrix | `cargo check -p <p> --all-targets --no-default-features [--features <set>] --message-format short` over 35 sets (none, each format, `sheet`, `markdown`, `eval`, `encryption`, and 16 combinations; `litchi-ops` skips the 6 sets that name a feature it lacks), run at the base and on this branch | `litchi-ops`: 29 of 29 pass, at the base and now. `litchi`: 17 pass and 18 fail, the same sets with the same error lines as at the base (see below). No new warnings. Two of the base's distinct warnings are gone: the unused probe-counter import in `detection_smart.rs`, and an unused variable in the deleted `document/doc.rs`. |
| Rustdoc | `RUSTDOCFLAGS=-Dwarnings cargo doc -p <p> --no-deps --all-features` for `litchi`, `litchi-ops` | clean, 2 of 2; `litchi` is also clean with `docx`, `docx,markdown` and `iwork` |
| Facade tests | `cargo test -p litchi --all-features --test <t>` for each target, then `--lib` and `--doc` | 30 targets, 180 tests; `--lib` 56; `--doc` 20 (1 ignored); all pass |
| Ops tests | `cargo test -p litchi-ops --all-features`; `cargo test -p litchi-ops --no-default-features --features <set> --lib` for `ods,xlsx`, `odp`, `docx`, `ppt` and `xlsx` | 388 unit and 57 doc tests pass; the subsets pass 106, 51, 126, 64 and 83 |
| Python oracle | `TMPDIR=<scratch> bash crates/pyo3-litchi/tests/run_tests.sh` | 132 tests OK |
| Crate boundaries | `python3 tools/check_crate_boundaries.py` | valid: 66 packages, 267 internal dependency declarations, 11 explicit debt items |
| Non-iWork gate | `python3 tools/non_iwork_gate.py verify` | 47 bulk roots, 35 facade-safe trees, 1 combined tree |
| Gate unit tests | `python3 -m unittest tools.test_check_crate_boundaries tools.test_non_iwork_gate` | 1,189 OK |
| Example targets | `python3 tools/check_example_targets.py`; `python3 -m unittest tools.test_check_example_targets tools.test_check_iwork_public_api` | exit 0; 22 OK |
| iWork public API | `python3 tools/check_iwork_public_api.py` | passes, and its rustdoc build is now warning-free |
| Harness clippy | `cargo clippy --manifest-path tools/perf-baseline/Cargo.toml --locked --offline --all-targets [--all-features] -- -D warnings` | clean, both ways |
| Harness tests | `cargo test --manifest-path tools/perf-baseline/Cargo.toml --locked --offline` | 632 passed, 1 ignored: lib 582, bins 45, `tests/unified_ops_filesystem` 1, `tests/xlsx_filesystem` 4 |
| Harness feature binaries | `cargo test --manifest-path tools/perf-baseline/Cargo.toml --locked --offline --features xlsb-crud,xls-source-attribution --bin xlsb_crud --bin xls_source_attribution` | 10 and 9 pass |
| Allocator target | `cargo check --manifest-path tools/perf-baseline/Cargo.toml --locked --offline --features allocator-metrics --bin litchi-perf-baseline-alloc` | builds |
| Source policy | `python3 -m unittest tools.test_perf_baseline_source_policy` | 28 OK |
| Coverage index | `python3 -m unittest tools.test_crud_coverage_index` | 50 OK |
| CI Python list | `python3 -m unittest tools.test_perf_compare tools.test_perf_smoke_baseline tools.test_perf_abba_summary tools.test_perf_abba_package tools.test_perf_resource_profile tools.test_perf_corpus_binding tools.test_validate_opc_serial_eager_allocator_abba tools.test_crud_coverage_index tools.test_validate_perf_default_matrix tools.test_perf_workflow_policy tools.test_perf_baseline_source_policy` | 517 OK |
| Index validator | `python3 tools/validate_crud_coverage_index.py --index docs/performance/crud-coverage-index-v2.json --catalog docs/performance/results/perf-corpus-manifest-v2.json --selector-source tools/perf-baseline/src/lib.rs --checklist docs/CRUD_Scenario_Checklist.md --repo-root .` | 15 categories, 33 mapped selectors |
| Claim registry | `python3 tools/check_perf_claims.py --registry docs/performance/claim-registry-v1.json --repo-root . --mode structural` | 10 claims OK. Strict mode and `tools.test_perf_claims` need evidence that this sparse worktree lacks, and `results/0251-xlsx-xml-borrowed-20260821/summary.json` is not in the tree at all. S9 changes none of their inputs. |
| Fuzz | `cargo check --manifest-path crates/litchi/fuzz/Cargo.toml --bins --keep-going` | 57 of 58 bins check. The exception is the known `keynote_chart_title` E0283 (iWork fuzz code, unchanged since `5533b7e3b3`), as at S8c. |
| Deletion proof | the script below | all 14 paths absent; both greps empty |
| `AGENTS.md` pre-commit checks | the two commands in `AGENTS.md` | both print nothing |

The 18 failing `litchi` sets in the feature matrix fail at the base in the same
way, on three targets that predate this step:

- `tests/tolerance_report_facade.rs` has no crate-level `cfg`, so it does not
  compile without a format feature. This accounts for 6 sets: none, `sheet`,
  `markdown`, `eval`, `encryption` and `sheet,eval`.
- `examples/comprehensive_docx_test.rs` requires only `docx`, but it imports
  `litchi::ooxml_common`, which needs the `ooxml-common` feature. This
  accounts for the 4 `docx` sets without `pptx` and `xlsx`.
- `examples/core_props_office.rs` requires `opc` and `ooxml-common`, but it
  imports `litchi::{docx, pptx, xlsx}`. This accounts for the 8 sets with
  `pptx`, `xlsx` or `xlsb` but not all three of `docx`, `pptx` and `xlsx`.

### Deletion proof

```sh
cd <worktree>
for path in crates/litchi/src/document crates/litchi/src/presentation \
  crates/litchi/src/sheet/workbook.rs crates/litchi/src/sheet/workbook_types.rs \
  crates/litchi/src/sheet/adapters crates/litchi/src/sheet/selection.rs \
  crates/litchi/src/sheet/functions.rs crates/litchi/src/markdown/document.rs \
  crates/litchi/src/markdown/docx.rs crates/litchi/src/markdown/presentation.rs \
  crates/litchi/src/markdown/writer.rs crates/litchi/tests/markdown_file_parity.rs \
  crates/litchi-ops/src/__migration.rs tools/perf-baseline/src/facade_ole2.rs; do
  if [ -e "$path" ]; then echo "PRESENT $path"; else echo "absent  $path"; fi
done
git grep -nE 'litchi::(Document|Presentation|Workbook|DocumentElement)\b|litchi::sheet::(Workbook|SelectedWorksheet|SelectedCellView|SelectedCell|WorksheetSelector|functions)\b|litchi::(document|presentation)::|litchi::sheet::open_|\bopen_(xls_|xlsb_|csv_|tsv_|prn_|text_|sylk_|dif_|fixed_width_)?workbook(_[a-z_]+)?\(|map_ooxml_error|__migration|crate::(document|presentation)::|facade_ole2' \
  -- 'crates/litchi/*.rs' 'crates/litchi/**/*.rs' 'crates/litchi-ops/**/*.rs' 'crates/pyo3-litchi/src/**' \
     'crates/pyo3-litchi/python/**' 'tools/perf-baseline/src/**' 'tools/perf-baseline/tests/**' \
  | grep -vE '^[^:]+:[0-9]+:\s*(//|\*)'
echo "grep exit=$?"
git grep -nE '^\s*pub (mod|use) (document|presentation)\b|pub use sheet::Workbook|pub use .*\bDocumentElement\b' -- crates/litchi/src/lib.rs
echo "grep exit=$?"
```

Output:

```text
absent  crates/litchi/src/document
absent  crates/litchi/src/presentation
absent  crates/litchi/src/sheet/workbook.rs
absent  crates/litchi/src/sheet/workbook_types.rs
absent  crates/litchi/src/sheet/adapters
absent  crates/litchi/src/sheet/selection.rs
absent  crates/litchi/src/sheet/functions.rs
absent  crates/litchi/src/markdown/document.rs
absent  crates/litchi/src/markdown/docx.rs
absent  crates/litchi/src/markdown/presentation.rs
absent  crates/litchi/src/markdown/writer.rs
absent  crates/litchi/tests/markdown_file_parity.rs
absent  crates/litchi-ops/src/__migration.rs
absent  tools/perf-baseline/src/facade_ole2.rs
grep exit=1
grep exit=1
```

Comment lines are excluded because several comments name the deletion. For
example, `litchi::Document` survives in the harness only in a comment in
`docx_source_selected_paragraph.rs` that records the change. The pathspecs do
match live code: `git grep -cE 'litchi::open|crate::open'` over them counts
matches in `detect/tail.rs`, `file.rs` and the format tests.

`git grep -n 'doc(hidden)' -- crates/litchi-ops crates/litchi/src` prints
nothing. At the base it printed 7 lines.

## 6. Notes for S10

1. **Selector names.** The post-S9 binary rejects the 47 removed names as
   `--case` values: `parse_case` returns `None` for them, and a guard test
   pins this. The A legs run them in the pre-M0a binary. The B group map uses the successor names
   in `results/0003/s8b/route-map.json`, as S8b's note 1 says. The 15 kept
   predecessors run in both binaries.
2. **`xlsb_crud`.** `full_text` now times `litchi::open_bytes` + `File::text`
   instead of the deleted `sheet::Workbook::from_bytes(..).text()`, under the
   same case name. Treat the pre-S9 and post-S9 results as a predecessor and
   its successor, not as one route. `open_identify` makes the same leaf call
   as before.
3. **`xls_source_attribution`** has no `facade-file` mode. Its recorded
   results keep that mode.
4. **Rustdoc without features.** `cargo doc -p litchi --no-deps` with no
   features, or with `markdown` alone, fails under `-Dwarnings` on the crate
   docs' links to `open`, `File`, `OpenOptions` and others, which need a
   format feature. This predates S9: S8d wrote those docs, and the gate builds
   with all features. It is a candidate for the M1 docs pass.
5. **XLSB coordinates** are unchanged. The source worksheets stay one-based,
   which is the M1 normalization item from S5c in the log. Their
   caller-coordinate classification should be decided together with that
   item.
6. **Errors outside the unified error.** These are left as they are, because
   none of them is in S9's scope:
   - The facade's `images` functions and `MetadataYaml` return
     `litchi::common::Error`. `MetadataYaml` flattens a YAML failure into
     `Error::Other` with a message.
   - The ops Markdown adapter reports a YAML front-matter failure as a
     message, because `Source` has no typed slot for `serde-saphyr`.
7. **Base movement.** `uo/m0a` moved to `ada71dbd50` (0014a and 0014b) after
   this branch's base. Those commits touch none of this branch's files.
8. `docs/LENIENCY_AUDIT.md`'s dated status (2026-10-02) still names the
   umbrella `Document`, `Presentation` and `Workbook`. It is a snapshot and
   was left as is.
