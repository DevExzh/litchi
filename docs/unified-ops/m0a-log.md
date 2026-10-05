# M0a step log (merged into PROGRESS.md when M0a lands)

## S2 — unified error (`7010055311`)
- Added `litchi_ops::error`: 39 codes (36 OfficeCLI + `unsupported_capability`, `limit_exceeded`, `cancelled`), a pointer-sized `Error`, a typed `Source`, and per-leaf classifiers for core/io/execution/opc/cfb/docx/pptx/xlsx/xlsb/doc/xls/ppt/rtf. 27 tests.
- 10 optional leaf edges registered (ADR 0024 amendment: 255 declarations / 244 edges). No iWork edges.
- Follow-ups:
  - (S6/S7) The leaf message is both the `Display` text and `source()`, so error-chain printers show it twice. Make `Display` facade-level only.
  - (M2) Add a diagnostics slot.
  - Add `Source::Crypto` with the `encryption` feature.
  - The design's code list has 36 entries, not the "37 + 1" that plan 0003 states.

## S3 — OpenOptions, Locked/Signed shells (`b9951ae15f`)
- `OpenOptions` (const `new()` = `Default`; common knobs plus feature-gated per-format limit sub-options; no credentials), `SourceMode`, `Family::of` (all 21 formats), and the pointer-sized `Locked`/`Signed` handles with redacted `Debug`. No unlock or unsign yet (S6/S6b).
- Merge with S2: the dependency and feature lists were combined. The only new edge is `litchi-ops -> litchi-odf-common`, which brings the inventory to 256 declarations and 245 canonical edges. The ADR 0024 amendments were reconciled. The boundary checker, the non-iWork gate and the litchi-ops clippy/tests pass on the merged tree.

## S4 — detection moved into `litchi_ops::detect` (`7b897d11a2`)
- 5,942 lines moved from `crates/litchi/src/detection_smart/`. Git records all five files as renames, and `ooxml.rs` is byte-identical. iWork detection stays in the facade as a lowest-precedence ZIP tail (`detect::tail::ZipTail`), so `litchi-ops` has no iWork edge and the bulk gate stays at 47. Facade call sites resolve through `#[doc(hidden)] __migration`, which S9 deletes. The new `test-support` feature in ops is enabled only through the facade's dev-dependency.
- Tests: 37 of the 45 in-source tests moved to ops, plus five new tail-ordering tests in ops and three Pages-precedence tests in the facade. The full facade suite passed on the S4 branch (554 passed, 11 ignored).
- Merge with S2/S3: the dependency and feature lists were combined into 15 ops leaf edges. Inventory: 262 declarations, 250 canonical edges. The ADR 0024 amendments were reconciled. The coordinator added a `clippy::large_enum_variant` allow on `PresentationSourcePathDetection`, which only fires in odp- or ppt-only builds; S4 had checked single-feature builds with `cargo check`, not clippy.
- Caveats:
  - `DetectedFormat` variants are now gated on litchi-ops features, so a second dependent that turns on extra ops features could break facade matches. This matters from S5 on.
  - Some moved doc comments still say the ops-only detectors classify iWork. Fix them in S9.
- Disk: S4's all-features facade test build reached 114 GB before it was cleaned. Building agents now run with `CARGO_INCREMENTAL=0` and debug info off.

## S5a — `Docx` handle (`5f7c0523f4`)
- `Docx(Arc<DocxState>)` holds a `Source`/`Eager` owner, lazy metadata, eager-route tolerance and the ported `finish` freshness guard. Constructors cover the source (path/bytes), eager (`OpcPackage`) and prepared routes; nothing is loaded at open. Typed reads: `text`, `metadata`, `tolerance_report`, `section_inventory`. A crate-internal `edit_document` builder works on both owners and refuses when `retain_source(false)`. 15 tests.
- `error.rs` gains `Source::DocxTransaction`. No leaf or dependency changes.
- Deferred: the Markdown materialization cache goes to S6/S7, since ops has no `markdown` feature. The paragraph/table/run `map_ooxml_error` sites go to S6's family views.
- **M2 blocker:** source-backed `edit_document()` refuses all 19 Word-authored fixtures with `UnsafeEdit` (markup-compatibility branch selection; every fixture declares `mc:Ignorable`). M2 needs a leaf fix, or a fallback that materializes the eager owner at first edit. Cross-check against the perf program's MCE rewrite (owner decision 2026-09-16 #1).

## S1 — pre-migration baseline (`6c96efff21`)
- All plan §5.3 routes measured at `da4c207478`: 52 child runs, 107 rows, pinned to one CPU, warm cache only. Results are in `docs/unified-ops/results/0003/before/` (README, analysis.json, summary.csv, raw/; `capture.py` and `analyze.py` can be reused for S10). Corpus hash `414fc710…8ba7d`.
- The pre-migration binary is kept at `/home/zhuhe/code/litchi-worktrees/uo-m0a-baseline-bin/litchi-perf-baseline-da4c207478` (sha256 `9c887ee5…2c79c04`) for the S10 ABBA comparison.
- Noise: the host was heavily shared (load 8–135 on 32 CPUs), so about a third of the rows exceed 5% spread. S10 rule: ABBA with at least 2 children per leg, median of per-child p50s, trigger = max(5%, route noise), a 0.5 µs floor for rows under 5 µs, and full separation required for rows above 10%. `doc_facade_file_open` is bimodal (~180 µs or ~246 µs per process).
- Gaps:
  - No per-route RSS (needs the allocator harness build, which was skipped to save disk).
  - `--corpus-manifest` fails on the detection group.
  - The plan's `docx_source_selected_paragraph` is a non-timing binary; `docx_semantic_one_paragraph` and `docx_producer_source_selected_paragraph` were measured instead.
- Disk: free space on `/` briefly reached 9.1 GB during the capture. The unified-ops worktrees were 2–3 GB each at the time, so most of that usage came from outside this program.

## S5e — tier-3 read handles Rtf/Odt/Ods/Odp (`ea0c856e3f`, after the rewrite)
- One-pointer `Send + Sync` handles over the facade's current payloads. Odt/Ods/Odp each have an `Eager`/`Source` owner; the `RefCell` around Ods is gone because every read API it uses takes `&self`. Constructors are `pub(crate)`, one per route. Errors are typed (`Source::Core` tagged with the format; `Source::Rtf`). `metadata()` returns `Option<Metadata>`, and `tolerance_report()` returns the leaf report type, which S6/S7 wraps. `chrono` is a new optional dependency on `rtf`. No leaf changes: all seven tier-3 leaf owners are `Send + Sync`.
- For S6:
  - The facade has no ODS sheet adapter and its `Workbook` does not implement `WorkbookTrait`; the `open_*_workbook_dyn*` openers refuse ODS. Default for S6: give `WorkbookView` the facade's existing ODS reads through `Ods::worksheets()`, and have the static `WorkbookTrait` path refuse ODS with `unsupported_capability`, as today. A real ODS adapter is new work for M1.
  - Use `Odp::from_prepared` instead of the facade's route, which turns a prepared ODP back into bytes and re-parses it.
  - ODF limits stay with detection.
- Observations:
  - `test-data/odf/corpus/impress-*.odp` never gets the source-backed owner from `detect_pptx_source_path_with_limits`: the path probe returns `Bytes`, so these files take the eager parse. The facade behaves the same today. It may affect the S10 ODP open timings.
  - ODP `metadata()` is `Some(Metadata::default())`, as in the facade, although the leaf can read `meta.xml`. This is an M1 fidelity item.
- 20 tests; clippy (all features and each single feature), docs, and the boundary/non-iWork gates pass. On the merged tree: litchi-ops clippy passes, and 115 unit tests plus 26 doctests pass.

## S5b — `Pptx` handle and L-PPTX-1 (`ff9fc9f0a9`, after the rewrite)
- `Pptx(Arc<PptxState>)` holds a `Source(SourceBackedPresentation)` or `Eager(Box<pptx::Package>)` owner with a lazy metadata cache. Constructors take detection output (source, source package, prepared, eager OPC). Typed reads: `text`, `slide_count`, `slides`, `slide`, `slide_catalog`, `slide_size`, `metadata`, `tolerance_report`. A changed source is `io_error` with the OPC `SourceChanged` kept as the leaf error (fixes F10). OPC read limits now report `limit_exceeded` and cancellation reports `cancelled` (the facade reported a core `ResourceLimit` and `Other`).
- L-PPTX-1: `SourceBackedPresentation::try_into_editor(self)` builds the editor from the package and slide list indexed at open, reading no source bytes. Open records one flag saying whether the slide list needed a repair. Leaf `Send + Sync` assertions (R3) pass.
- 13 ops tests and 5 leaf tests. Two `source_backed_svg_lifecycle` tests need the untracked `3rdparty` corpus that the sparse worktree lacks; they pass with the corpus present.

## S5c — `Xlsx`/`Xlsb` handles, L-XLSX-1, L-XLSB-1 (`b32505fc8d`, after the rewrite)
- Two-state `Arc` handles. The facade's sheet adapters and selection moved over with typed errors. `WorkbookTrait` is implemented directly on both handles, so S6 delegates by `match`. Case-folding is unchanged: case-insensitive for XLSX, exact for XLSB (F15). `Xlsx::sheet(selector)` is the leak method.
- Leaf changes:
  - L-XLSX-1: `SourceBackedWorkbook::source_backed_package()` and `try_into_source_backed_package()`.
  - L-XLSB-1: `SourceBackedWorkbook::to_workbook()`.
  - R3 `Send + Sync` assertions pass.
- 24 ops tests, including route parity against values captured from the current facade.
- Four litchi-xlsx test targets (`form_control_properties`, `form_control_scalar_lifecycle`, `form_control_owner_read`, `drawing_svg_read`) `include_bytes!` files that the sparse checkout omits. Run them on a full checkout before the M0a merge.
- For S6:
  - Declare `format/cell_text.rs` once in `format/mod.rs`; it is currently `#[path]`-included by both handles.
  - XLSB source worksheets keep one-based coordinates while the eager owner is zero-based, as in the facade. This is an M1 normalization item.

## Cross-cutting M2 prerequisite (R4): editing a shared snapshot
- The leaf source-backed publishers consume their package: `docx` `publish_document_commit_to_stream(self)`, the pptx `SourceBackedPresentationEditor` publishers, and the xlsx `SourceBackedPackage`. None of these packages can be cloned.
- So a handle can hand its package to an editor only when it is the sole owner. Per format, when the handle is shared:
  - xlsx falls back to an eager snapshot built from the existing index (no re-index);
  - pptx has no fallback, because the reader is not retained;
  - docx has the same publish-time constraint.
- ADR 0003 and design §3/§8 require the old snapshot to stay readable after a write. M2 therefore needs one of: a non-consuming publish (or cheap package clone) in the leaves or litchi-opc; a retained reader for a cheap re-open; or consuming Rust verbs plus a documented clone-then-edit fallback. Record 0009 will decide this before M2.
- Merged tree after S5a/b/c/e: litchi-ops clippy passes with all features and with each of pptx, xlsx, xlsb, xlsx+xlsb and pptx+xlsx; 153 unit tests and 29 doctests pass; `cargo check -p litchi --all-features` passes.

## S5d — legacy OLE handles `Doc`/`Xls`/`Ppt`, L-DOC-1, L-PPT-1, L-XLS-1 (`2bec864ec1`, after the rewrite)
- One-pointer `Send + Sync` handles. Each retains its source without a copy or a second parse:
  - The doc and ppt byte routes move detection's in-memory CFB parse into a `SharedOleFile` (`SharedOleFile::from_owned_ole_file`).
  - The ppt path route keeps the `SourceBackedPackage` that the facade used to drop.
  - `retain_source(false)` releases the source.
- `Xls` has two states: `Source` (source-backed workbook plus `SharedOleFile`) and `Eager`. `XlsFile` and `Other` (F12) are not carried over.
- Metadata keeps the facade's return shapes; S6 unifies them. The editor builders return leaf-typed errors until M2.
- Leaf changes:
  - `litchi-cfb`: `SharedOleFile::from_owned_ole_file` and `is_owned_immutable`.
  - `litchi-ole-common`: `SourceBackedOverlayPublisher::from_shared`.
  - L-DOC-1: `body_text::source::SourceSnapshot::from_shared`, plus `Package::into_ole_file`.
  - L-PPT-1: `text_edit::SourceSnapshot::from_shared`, `SourceBackedPackage::text_edit_snapshot[_with_limits]`, and `Package::into_ole_file`.
  - L-XLS-1 uses the plan's fallback: `cell_values::Snapshot::from_shared_ole_file` copies the source once and indexes it at the first edit. The editor captures streams through the cursor-based reader, so it cannot avoid the re-index. The copy cost is about 23 µs vs 21 µs from caller bytes on `Simple.xls`.
  - R3 `Send + Sync` assertions in the leaves.
- New edge `litchi-ops -> litchi-ole-common` (doc/ppt/xls features). Inventory: 263 declarations, 251 canonical edges.
- **Deviation, provisional default: the `.xls` byte route stays eager.** Over the 54 corpus workbooks:
  - the source-backed owner refuses BIFF5 and older files;
  - it fails text extraction on 4 files the eager reader reads;
  - the eager reader omits worksheets in 8 files the source owner reads in full.

  The facade's two routes already disagree on those 8 files today, and a test pins the list. Keeping today's routing means no regression in M0a. Reconciling the two readers is a correctness item for the xls write stream (M2b), because a byte-opened and a path-opened `.xls` currently show different sheets.
- Quick open check (release, 200 iterations; not the S10 protocol): `SampleShow.ppt` 31.8 → 31.7 µs, `testPictures.doc` 48.5–49.5 → 48.8–49.0 µs. Retention (R5) keeps exactly the input bytes alive; S10 measures RSS.
- Already failing before this step, not touched: `cargo doc -p litchi-doc` fails on private intra-doc links in `parts/{bookmarks,comments,rsids,sections/codec}.rs`.
- Merged tree after all S5 steps:
  - litchi-ops clippy passes with all features and with each of doc, ppt, xls, pptx, doc+ppt+xls and docx+doc.
  - 173 unit tests and 29 doctests pass, and `cargo check -p litchi --all-features` passes.
  - The boundary checker (263 declarations), the non-iWork gate and the gate unit tests pass.

## Coordinator check — S5c's fixture-dependent xlsx tests
- On `uo/m0a` (S5c merged), with the evidence directory `docs/report/spec-gap-validation-evidence/xlsx-form-control-properties/` added to the sparse checkout and `3rdparty` temporarily symlinked from the main checkout, `cargo test -p litchi-xlsx --all-features` passes the four targets S5c could not compile: `form_control_properties` (15), `form_control_scalar_lifecycle` (39), `form_control_owner_read` (18) and `drawing_svg_read` (32). This closes S5c's open item.

## S8a-pre — Python read-API oracle (`70f8465a72`)
- 121 `unittest` tests in `crates/pyo3-litchi/tests/` (14 expected failures) pin today's `litchi_py` read API:
  - docx/doc/rtf/odt/pages through `Document`; pptx/odp/key through `Presentation`; xlsx/xlsb/ods/numbers through `Workbook`;
  - detection, errors, and `.pyi`/README sync.
- Green on dev and release builds, and under pytest 9.1.1 (in a temporary venv, since deleted). Runner: `crates/pyo3-litchi/tests/run_tests.sh` (cargo build, no maturin). Record: `docs/unified-ops/s8a-python-oracle.md`.
- Drift pinned as today/intended test pairs:
  - D1: `.odb` is detected as `Doc`.
  - D2: the stub lists 14 of 21 formats.
  - D3: the stub calls `FileFormat` an `Enum`; it is not.
  - D4: exception classes contradict the stub's `Raises:` notes.
  - D5: the README documents a `Worksheet` API that does not exist.
  - F10: PPTX raises `ValueError` on a source change; the other formats raise a bare `Exception`.
- Bugs found, to fix in S8a:
  - B1: `.ppt`/`.xls` are unrecognised; the pyo3 crate enables neither the `ppt` nor the `xls` feature, although its docs claim both.
  - B2: `repr()` panics when its byte-sliced preview cuts a multibyte character. That is an interpreter abort in release builds (`drawing.docx`).
  - B3: DOCX `Document.text()` joins paragraphs with no separator.
  - B4: `Run.underline()` is a placeholder that always returns `None`.
  - B5: using a handle from another thread panics. The unified `File` is `Send + Sync`, so S8a can make the classes thread-safe.
- S8a contract:
  - `Document`/`Presentation` I/O errors stay `OSError`.
  - Flat ODF, password-protected `.doc` and Word 6 stay `ValueError`, including once a protected `.doc` becomes `File::Locked`.
  - Wrong-opener messages keep naming the detected format.
  - Exact messages and `repr` snapshots may change on purpose, and the S8a record lists each change.

## 0012 — litchi-opc `SourceBackedPackage: Clone` (0009 stage S0; `0cf1f43403`, merged here)
- The archive index, the admitted catalog (8 fields grouped into `AdmittedCatalog`) and the cache are behind `Arc`. `Clone` is hand-written and lists every field, and a compile-time check keeps the type `Clone + Send + Sync`.
- A clone is the same opened package: same lineage, version, cache, read-ahead window and budget. A clone costs 6 reference counts (11–12 for a package opened with an execution budget) and no allocation. Publisher signatures are unchanged; each publication consumes a clone.
- Two behaviour changes:
  - `into_opc_package` now copies the non-part member list, so it can fail with the same `Allocation` error as `to_opc_package`.
  - The splice-fragment and read-session identity checks compare the opened package instead of the instance address.
- 18 new litchi-opc tests: 0009's T1–T3 and T5–T8 at the OPC level, budget accounting, concurrent compression waves, and a lineage proof across clones. Leaf suites: opc 1,042, docx 2,010, pptx 1,305, xlsx 2,133, xlsb 1,030 passed.
- Quick ABBA loop (release build, pinned CPU, shared host): open and full-text ratios 0.977–1.003, so no slowdown. Each open makes exactly 3 more allocations; clones and reads make none. Clone plus drop takes 72–77 ns (122–130 ns with a budget), against an 80 µs open. Evidence and the benchmark source are in `docs/unified-ops/results/0012/`; binaries are not committed.

## S6 — `File`, single-pass open, family views, Locked/Signed (`a8d3ce9772`, `eca894f849`)
- `File` is flat, `#[non_exhaustive]` and 16 bytes, with one variant per format plus `Locked`/`Signed`.
- `open`, `open_with`, `open_bytes`, `open_bytes_with` and `open_prepared[_with]` each build one OPC index per open, proven by the probe counter. `#[doc(hidden)] open_with_tail` / `open_bytes_with_tail` let S7 plug in the iWork tail.
- `File::{format, family, text, metadata, tolerance_report, document, presentation, workbook}`. `ToleranceReport` moved into ops (with `Untracked(FileFormat)`), and the compat wrappers moved to `compat::{word, slide}`. `WorkbookView: WorkbookTrait`; ODS reports names and count but refuses worksheet access through the trait.
- Locked/Signed:
  - `Locked(EncryptionKind)` for encrypted DOC, XLS and PPT and for OOXML in a compound file.
  - `unlock(Password)` under the new `encryption` feature (edge `litchi-ops -> litchi-crypto`; 264 declarations). Every legacy fixture round-trips by path and by bytes; encrypted DOCX/XLSX/PPTX are built in-test and unlock to the plain text.
  - `Signed(Opc|Cfb)` wraps the readable inner file and delegates reads; it stays read-only until S6b.
- Record 0008 findings:
  - A1: an execution context is only checked for cancellation around detection. No open builds the managed DOCX owner, and all 19 Word fixtures read under a context.
  - A5: `SourceMode::Eager` uses the eager OPC route, keeping the ZIP layout, except `open_prepared` under Eager.
- Deviations:
  - `File::format()` and `Locked::format()` return `Option`: an encrypted OOXML package opened from bytes has no knowable format.
  - `Error` is `Clone`, so `into_leaf` is gone; use `leaf()`.
  - The compat enums gained `#[non_exhaustive]` and lost their Pages/Keynote variants, since ops is iWork-free.
  - The S5 `from_prepared` constructors are removed.
  - XLS reads through `WorkbookTrait` come from an eager reader, loaded on first trait use.
  - Disabled-feature hosts are reported with their format and a feature suggestion; the facade silently fell back to XLS. S7 keeps the new behaviour.
- Open cost (release, 300 interleaved rounds, median, facade → ops):
  - Within noise for docx, doc, pptx, xls and xlsx on both routes, and for ppt by path.
  - PPT bytes: 50.9 → 53.7 µs (+2.8 µs). That is S5d's default source retention; with `retain_source(false)` it is 50.9 µs. It is within the goal's "5% or 20 µs" budget.
- Verified on the merged `uo/m0a` (S6 + 0012):
  - litchi-ops clippy passes with all features, with no default features, and across 16 single and combined feature sets.
  - 360 unit tests and 52 doctests pass, and `cargo check -p litchi --all-features` passes.
  - The boundary checker (264 declarations), the non-iWork gate and the gate unit tests pass.
- Left for S7: the Markdown adapter and its tests; the facade reads slide size and catalog from the `Pptx` handle.
