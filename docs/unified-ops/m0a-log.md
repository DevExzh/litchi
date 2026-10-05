# M0a step log (landed on trunk in `1927763795`; summary in PROGRESS.md)

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

## 0011 — DOCX MCE admission, stages 1–3 of 0008 (`53f938e83c`, merged here)
- A shared single-pass check in `litchi_ooxml_common::mce::admission` decides, without building the processed copy, whether MCE processing could hide anything a reader sees. It reuses the codec's name resolution and limits, and admits a part only when the codec would also accept it (a generated test checks 680 inputs). When the check cannot admit a part, the old code path runs, so refused documents keep today's errors.
- Escaped namespace URIs are now decided on their decoded value in the codec, its fast path and `active_offsets`. About 20 xlsx/pptx/DrawingML/docx scanners that read raw namespace values now refuse such parts; regression tests were added.
- Fixtures now accepted, out of 19:

  | Route | Before | After |
  |---|---|---|
  | Source-backed `edit_document` | 0 | 19 |
  | Source-backed changed publish | 0 | 16 (the 3 protected files stay refused) |
  | Managed read and snapshot | 0 | 19 |
  | Document variables | 15 | 19 |

  Over the 63-file corpus, source-backed edits went from 1 to 59. The M0a S5a refusal test now expects success, and an `AlternateContent` refusal test replaces it.
- Breaking changes to the eager owner:
  - It applies the same admission rule.
  - Changed publication refuses enforced protection and tracked revisions, with the source-backed route's exact error. That covers the 3 protected fixtures and `redline-range-comment.docx`.
  - Paragraph transfer refuses donors whose MCE processing drops markup (`TransferRefusal::MarkupCompatibilityScope`).
  - A nested `mc:Ignorable` listing only understood namespaces is admitted, because Litchi's own revision writer emits `w16du`.
- Indicative performance: the check costs 0.26–0.41× of a full MCE pass. Source-backed open plus `edit_document` costs 0.57–0.91× of the old refusal. The no-MCE path is unchanged (0.96–1.00×). Eager open plus edit is 1.07–1.84× slower, because it now runs the admission and protection checks it never ran.
- Verified on the merged `uo/m0a`:
  - fmt and clippy `-D warnings` on ops, docx, ooxml-common and DrawingML;
  - litchi-ops: 361 unit tests and 52 doctests;
  - full suites (with `3rdparty` symlinked temporarily): docx 2,022, ooxml-common 451, DrawingML 369, pptx 1,309, xlsx 2,137, all passing.

## S7 — facade switch (`28962cfe88`)
- The facade root re-exports the ops API: `open*`, `File`, `OpenOptions`, `SourceMode`, `Error`, `Result`, `Code`, `Locked`, `Signed`, `Family`, `ToleranceReport`, the views, `EncryptionKind`, `SignatureOrigin`, `DocOptions`, `open_prepared*`, `Password` (encryption), and the modules `error`, `format` and `compat`.
- `litchi::open*` (`crates/litchi/src/unified.rs`) plug in the iWork tail. Pages, Keynote and Numbers packages and package directories return `iwork_unsupported` with the detected format and a `litchi::iwork::Document::open` suggestion, tested on path, eager and bytes routes.
- `litchi::Error`/`Result` are now the ops types. The old `Document`/`Presentation`/`Workbook` keep `litchi_core::Error` as `litchi::common::Error` until S9. Moving them would touch about 10k lines that S9 deletes anyway.
- Markdown:
  - `impl ToMarkdown for File` lives in litchi-ops because of Rust's orphan rule, behind new ops features `markdown`, `formula` and `yaml`. New edges `litchi-ops -> litchi-markdown` and `litchi-ops -> litchi-formula` bring the inventory to 266 declarations and 254 canonical edges.
  - `ToMarkdown` gained `type Error`, a breaking trait change that avoids string flattening, plus a new `Source::Formula`.
  - The DOCX Markdown cache moved onto the `Docx` handle, keeping its freshness checks.
- Parity: `File` Markdown against the old types over 438 fixtures × 3 option sets × 2 routes: 2,076 identical, 552 refused by both, 0 differing. Tolerance reports are equal on every `tolerance_report_facade` fixture.
- Known: some smaller feature sets fail the facade's `--all-targets` clippy run, but with identical error lists on the base commit (checked on 10 feature sets). The iWork feature sets fail on a lint inside `litchi-iwa-archive`, which this program does not edit.
- For S9:
  - delete the facade's old Markdown adapter with the old types;
  - keep a public `ZipTail` path when `__migration` is removed;
  - point `litchi::SlideDescriptor` at the ops type.
- Verified on the merged `uo/m0a`:
  - fmt, plus clippy `-D warnings` on litchi and litchi-ops with all features;
  - litchi-ops: 442 tests;
  - facade targets: `markdown_file_parity` (9), `unified_open_facade` (14), `tolerance_report_facade` (12), `unexpected_format` (16), `read_limits_facade` (8), the three Markdown suites, and `--lib` (352);
  - the boundary checker, the non-iWork gate and the gate unit tests.

## S6b — signature policy (`f30c521c90`, record 0013)
- New API:
  - `Signed::unsign(self) -> File` (O(1), `#[must_use]`) and `File::is_signed()`.
  - Tier 1–2 handles carry a signature state (`Arc<Opened<State>>`, still one pointer; `File` is still 16 bytes).
- Editor builders:
  - While a file is signed, all seven refuse with `document_protected`, naming `Signed::unsign()`.
  - After `unsign` they build over a copy whose signature the leaf removed explicitly. For OOXML that is the leaf's `unsign` on a package opened over the exact input bytes; for DOC, XLS and PPT it is `litchi_sign::cfb::Editor::clear`.
  - An unlocked encrypted DOCX or PPTX refuses rather than publish in the clear.
- `resign` is deferred to P1: the unified layer publishes nothing until M2, and signer input must meet ADR 0006's credential rules.
- New edge `litchi-ops -> litchi-sign`.
- Leaf fixes (breaking):
  - pptx `opened::Transaction::commit` and about ten xlsb families stripped signatures silently; they now refuse.
  - Added: `pptx::Package::unsign`, `xlsx::Workbook::unsigned`, and read-only accessors (`source_backed_package()`, `shared_ole_file()`).
  - Exact no-ops on signed input stay byte-identical.
  - Five tests that pinned the old stripping now assert the refusal.
- Audit test: 11 signed inputs × 4 routes (3 DOCX, 2 PPTX and 2 XLSX from the corpus; XLSB, DOC, XLS and PPT signed in-test), plus the encrypted-and-signed refusal.
- Follow-ups:
  - About 115 package-level implicit `unsign()` calls in pptx (~60), xlsx (~45) and docx (~10) still strip outside the builders' reach.
  - Tier 3: the ODF rebuild drops signature members (`litchi-odf-common/src/core/writer.rs:1621`), reached from the ODF formula, ODP and ODS mutators.
  - The opener does not detect ODF signatures.
  - xlsb task panes and ribbon refuse only at save.
- Merge into `uo/m0a` (after S7): I combined the appended classifiers and tests in `error.rs`, the `Docx` state fields (the S7 Markdown cache and the S6b `ooxml_limits`), and the ADR 0024 amendments. Inventory: 267 declarations, 255 canonical edges.
- Verified on the merged tree:
  - litchi-ops/litchi clippy with all features, and litchi-ops clippy across 10 feature sets;
  - litchi-ops: 449 tests, including the 26 signature tests;
  - facade targets: parity, open, tolerance, unexpected-format, read-limits and encryption;
  - leaf suites: xlsb 1,039, pptx 1,314, ppt 1,349;
  - the boundary checker (267), the non-iWork gate and the gate unit tests.

## S8d — examples and docs on `litchi::open` (`7e6746e2a3`)
- 16 examples ported to `litchi::open`/`File`: the `sheet_eval_*` examples use `FormulaEvaluator::new(&file.workbook()?)`, statically dispatched. `text_formats_api_demo`, missed by plan §5.2, moved to the concrete `TextWorkbook`. The CLI examples print the error's cause chain.
- Rewritten: the README Quick Start (with a "not done yet" list), the guide's Unified API section, the crate rustdoc, and the `ToMarkdown` doctests.
- All 16 doc snippets were compiled and run. Example output before and after is identical except `NOW()`, iWork rows (now plain text) and the unified error text.
- Stale README snippets fixed: the XLSX writer, the formula snippet, and the iWork-directory note.
- Findings, for the M0a close-out (S9):
  - XLSX coordinate and range errors (`cell("A0")`) classify as `corrupt_file` through the wildcard arm. They are caller errors and should be `invalid_argument` or `out_of_range`.
  - `Docx::section_inventory` refuses Word-authored files (MCE bindings, F-15 in record 0010); L-DOCX-R5 is asked to cover it.
- Verified on the merged `uo/m0a`: facade examples clippy, 52 doctests, and the example-targets gate.

## S8c — facade tests on `litchi::open` and `File` (`940f8fb440`)
- All facade integration tests and the in-source `prepared_index_tests` now use `open*`, `File` and the views: 192 tests over 31 targets, plus `--lib` 352 and `--doc` 49. Only `markdown_file_parity` still names the old types, by design.
- Eval: `FormulaEvaluator` reads `WorkbookView` directly, which closes S6's item. A new test evaluates through the path, bytes and eager routes.
- S7's parity checks moved to 47 KB of goldens in `crates/litchi/tests/golden/`, recorded from the old types:
  - Markdown digests for the 438-file corpus;
  - 25 tolerance reports;
  - 55 PPTX slide catalogs.

  `markdown_file_parity` pins the goldens to the old types until S9 deletes that file. `LITCHI_BLESS_GOLDEN=1` regenerates them.
- Six assertions changed form or could not be ported (record `docs/unified-ops/s8c-facade-tests.md` §3). For example, the "XLSB has no sheet selector" refusal is now enforced at compile time, and the content-types limit now names `ContentTypesBytes` instead of the flattened `InputBytes`.
- Fuzz: no target uses the old types. 57 of 58 bins check; `keynote_chart_title` has an E0283 that predates this work (iWork, not edited).
- For S9:
  - Delete `markdown_file_parity.rs` whole.
  - Keep the detection, `sheet` and `markdown` re-exports the tests use.
  - Give the OPC probe counter a public test path when `__migration` is removed.
  - The full list is in §5 of the record.
- **Merge note for `uo/m1`:** L-PPT-R1 (already on `uo/m1`) changes PPT slide text, so the Markdown golden's PPT rows will differ there. Re-bless them on `uo/m1` after this merges, and record the reason (the F-10 fix).
- Verified on the merged `uo/m0a`: facade clippy with all features, plus `markdown_corpus_golden`, `markdown_file_parity` (12), `unified_open_facade`, `unexpected_format`, `read_limits_facade`, `xlsb_facade` (23), `eval_xlsx_integration`, `unified_sheet_selectors`, `tolerance_report_facade` and `pages_facade`.

## S8b — perf-baseline unified and leaf routes (`d5469842fd`)
- 118 opt-in selectors: `<prefix>_unified_<rest>` (through `litchi::open*`, `File`, views and handles) and `<prefix>_leaf_<rest>` (the same leaf calls without detection), for every plan §5.3 scenario. Old routes stay until S9.
- Unpaired:
  - ODS selected cell and media: `File` has no ODS cell or media read in M0a.
  - The untimed docx binary.
  - ODF polyglot and catalog-alias detection have no leaf leg.
  - Two DOCX read-one successors time a wider scope; S10 compares them only against their leaf twins.
- Registry 534 → 652, and the CRUD index and its test are bound to 652, which fixes F-11. Also fixed: the harness's stale `Cargo.lock` (broken since S4) and a clippy failure that predates this work.
- Route map and smoke summary are in `docs/unified-ops/results/0003/s8b/`.
- Gates: fmt, clippy `-D warnings`, `cargo test` (643 passed, 1 ignored), the Python policy suites (28 + 50) and the CI Python list (517 tests). One release build ran all 118 routes once, with no measurement kept.
- Observation: unified errors print only the code summary; the specific cause is reachable only through `Error::leaf()`. So legs refusing the same file record differently worded messages (`duplicate-style-names.doc`).
- The coordinator merged this into `uo/m0a` without a separate harness re-run: the harness is a standalone workspace, and S8b's own gates ran on this base.

## S8a — Python bindings on `litchi::File` (`e248da4eac`)
- Python `Document`, `Presentation` and `Workbook` open through `litchi::open` and read through the family views and `compat` wrappers.
- iWork reads go through `litchi::{pages,keynote,numbers}::semantic::Document`. The `litchi::iwork` snapshot lacks Numbers sheet names, number and boolean cells in Numbers text, and the Keynote show title, all of which the oracle pins (R6). Output equals the old facade on all 123 iWork fixtures, and iWork package directories now open.
- New exception hierarchy (names provisional for M3): `LitchiError` with `PathError` (also `LookupError`), `PropertyError`, `UnsupportedError`, `CorruptFileError` and `ProtectedError` (also `ValueError`), and `IoError` (also `OSError`). Each carries `code`, `message` and `suggestion`. The oracle contract holds.
- Fixed:
  - B1: `.ppt` and `.xls` now open (the features were enabled).
  - B2: `repr` cuts at character boundaries.
  - B3: DOCX text has one line per paragraph, via a new validating paragraph mode in the litchi-docx text scanner (`text_with_paragraph_separator`). This changes `File::text()` for DOCX in Rust too. It costs +1.8% instructions on that path; the plain path is unchanged.
  - B4: `underline()` returns real values (new `compat::word::Run::underline`).
  - B5: handles are thread-safe.
  - D1–D5 (`FileFormat.Odb` = 20, stub and README) and F10.
- Oracle: 121 tests with 14 expected failures became 132 tests, all passing, in dev and release. The fixed expected failures became ordinary tests, their "today" twins were deleted, and 24 tests were added.
- For S9/S10:
  - Python iWork opens are 4–6% slower (+65–73 µs), because `litchi::open` scans before refusing iWork and then the iWork reader detects again. This lasts until M6, or until the facade hands over the scanned source.
  - DOCX full-text timings taken before B3 are not comparable. The S8b leaf leg must also use `text_with_paragraph_separator` for a fair pair.
  - Not fixed: a CryptoAPI-encrypted `.ppt` classifies as `unsupported_capability` instead of `document_protected`.
- Verified on the merged `uo/m0a`:
  - clippy on litchi-py, ops, litchi and docx with all features;
  - litchi-ops: 452 tests;
  - facade targets: `unified_open_facade`, the Markdown goldens and parity, tolerance, unexpected-format, semantic, and the Pages/Keynote/Numbers suites;
  - Python oracle: 132 OK.

## 0014b — xlsx/docx implicit `unsign()` (0013 F1) and XLSB task panes/Ribbon (F2) (`fe2c26f068`, merged here)
- Classified every site:
  - xlsx: 49 sites. 36 stripped silently on a changed edit, and 6 rewrote identical content (the slicer, timeline and query-table reorders and the smart-tag store); no site followed a refusal. The other 7 are the explicit authorization (`Workbook::unsigned`).
  - docx: 13 sites. 11 stripped silently; 2 are the explicit authorization (`Package::unsign`).
- Changed edits on signed packages now refuse with the crate's typed signed error before mutating. Exact no-ops keep the signature byte for byte. New `xlsx::Package::unsign()` is the explicit path for the slicer, timeline and smart-tag editors.
- Breaking changes:
  - xlsx changes on signed packages return `Error::Signed`; connections, XML Maps and volatile dependencies return it boxed.
  - Four xlsx paths now refuse before touching the package instead of failing at save.
  - docx mail merge, glossary, web settings, the font table and font settings refuse with `Opc(SignedSourceRequiresExplicitPolicy)` (reported as `document_protected`).
  - docx `put_ribbon`/`remove_ribbon` refuse up front.
  - Reorders that produce identical content now leave the bytes untouched on unsigned packages too.
- F2 (XLSB): task panes now return a typed `PackageError::Signed` (they used to return an untyped error that ops reported as `corrupt_file`), and the Ribbon refuses instead of stripping. The record corrects record 0013 §2.1 accordingly.
- Follow-ups:
  - F1b: litchi-ooxml-common's shared custom/core-properties, Custom XML, VBA and Ribbon code still strips when called directly from xlsx, docx, pptx and xlsb. One record together with 0014a.
  - F1c: semantic no-ops through a different serializer are still refused.
  - An unrelated bug: adding a second timeline cache fails once a timeline view exists.
- Verified on the merged `uo/m0a`: lib clippy on xlsx, docx, xlsb and ops; the new signature-policy targets (xlsx 16, docx 7, xlsb 10); the ops signature audit (26).

## 0014a — pptx implicit `unsign()` (0013 F1) and the shared litchi-ooxml-common owners (F1b) (`88729fa29f`, `02123283f1`, merged here)
- litchi-pptx, 63 sites:
  - 2 kept: they follow the caller's explicit `PlaceholderSignaturePolicy::Invalidate`;
  - 53 silent strips (one test-only) now refuse with `Error::Opc(SignedSourceRequiresExplicitPolicy)` until `Package::unsign()`;
  - 8 that also stripped on a no-op now write nothing when the result is unchanged (comment, author and modern-comment commits, theme `put_*`).
- Also in pptx:
  - Shape classification, design elements and Designer properties now detect a no-op before taking a mutable borrow.
  - `Package::edit_opc` refuses a changing closure on a signed package.
  - Opened-transaction staging verbs refuse at staging instead of at commit.
  - `store_placeholder_shape` refuses instead of always stripping.
  - The unsigned output of 94 operations is byte-identical before and after.
- F1b, litchi-ooxml-common: the 13 silent strips (custom/core properties, Custom XML, VBA, Ribbon) now refuse. Each caller keeps its own signed error: pptx and docx map it to `Error::Opc`, xlsx to `Error::Signed`. A docx managed save with edited core or custom properties rolls back and keeps the edit pending.
- Tests: 10 new in pptx (a 27-family table on both signed fixtures) plus 1 unit test; 3 in ooxml-common; 2 in docx; 3 in xlsx. 19 tests that asserted the old stripping were rewritten.
- Follow-ups:
  - F1-late: some pptx stores (`store_presentation_comments`, `store_modern_comment`, `media_parts::store`, `slide_sync::store`) only fail at save instead of refusing up front. Nothing is stripped.
  - F1-semantic: no-ops through a different serializer are refused rather than treated as no-ops.
  - Unrelated: shape classification and design elements read one shape on its own, so they fail on producer slides that declare `p:` on the root.
- Verified on the merged `uo/m0a`: lib clippy on pptx, ooxml-common, docx, xlsx and ops; pptx `package_signature_policy` (10); ooxml-common (454); litchi-ops (452).

## S9 — format-family types and `__migration` deleted (`59bc8f8772`, `27e35692bf`; record `docs/unified-ops/s9-deletion.md`)
- Deleted from the facade (24 files, 18,495 lines):
  - `Document`, `Presentation` and `sheet::Workbook`, with their element types;
  - the 30 sheet openers (the plan's "42" was wrong);
  - `map_ooxml_error` (114 call lines);
  - the old Markdown adapter and `markdown_file_parity`;
  - the facade's `base64`, `chrono` and `memchr` dependencies.
- Also deleted:
  - in ops, `__migration` and the detectors only it exported (`detected.rs` 4,261 → 882 lines); no `#[doc(hidden)]` remains;
  - in the harness, `facade_ole2.rs` and 47 predecessor selectors (registry 652 → 605, CRUD index and policy tests re-bound).
- In all, 76 files changed: +1,803 and −27,382 lines.
- Close-out fixes:
  - XLSX coordinates and ranges from the caller are `invalid_argument` or `out_of_range`; references read from the file stay `corrupt_file`.
  - A CryptoAPI `.ppt` opens as `File::Locked` in builds without `encryption`.
  - The S4 iWork doc comments are fixed.
  - `ZipTail`, the `_with_tail` functions and the probe counter have public documented paths.
  - `litchi::SlideDescriptor` is the ops type.
  - `Display` was already facade-level.
- Kept:
  - `litchi::common::Error` (= `litchi_core::Error`, the typed `Source::Core` leaf);
  - `sheet::text`, `iwork`, the leaf modules, and the re-exports the tests use;
  - 15 predecessor selectors that never timed the deleted types, plus the ODS selected-cell/media routes (no successor until M1).
- Gates: clippy and rustdoc clean; 30 facade targets (180 tests); ops 388 + 57; Python 132; harness 632; policy tests 28/50/517; boundaries; non-iWork 47.
  - The failures that remain match the base exactly: workspace fmt (iwa-archive, soapberry-zip), 18 facade feature-subset sets with the same errors, the `keynote_chart_title` fuzz target, and rustdoc without features.
- Merged into `uo/m0a` after 0014a/b:
  - no conflicts;
  - clippy (litchi, ops, litchi-py, all features); ops 445 tests;
  - facade targets: unified open, Markdown golden, tolerance, unexpected format, read limits, xlsb, eval, sheet selectors, encryption, pages; lib 56;
  - Python 132;
  - the boundary and non-iWork gates and their tests, and the perf policy and CRUD index tests.
- Note: `tools.test_perf_baseline_source_policy` and `tools.test_crud_coverage_index` read top-level files under `docs/performance/results/`, which the sparse worktrees exclude. With `/docs/performance/results/*.json` and `*.md` added to the sparse checkout, both pass on the merged `uo/m0a`.

## S10 — before/after performance gate (`d10dd4c747`, `c475a527df`, `1c00c6cd68`; record `docs/unified-ops/s10-measurement.md`)
- Verdict: M0a passes the gate on the plan's terms ("blocks the merge until explained or fixed"), not under the raw rule.
  - 91 of the 99 facade rows pass S1's rule, and all 47 kept predecessors pass (0.87–1.03).
  - The other 8 are explained by two measured mechanisms. Neither is a cost of M0a's code, and no fix to M0a's code is needed.
- E1, the first call in a fresh process (7 rows): XLSX open and open lifecycle, both PPTX slide counts, PPTX selected slide, and DOCX eager paragraph count and full text.
  - These routes time one call in a fresh child, so their result depends on the harness executable.
  - Standalone probes built from the same sources give the same first-call time at both revisions: XLSX open 187.6 vs 187.9 µs, PPTX `slide_count` 0.54 vs 0.47 µs.
  - Callgrind shows the same work (XLSX open +0.75% instructions).
- E2, the RTF sample loop's allocator state (`rtf_file_unified_open`): the untimed projection decides whether the next timed open re-faults about 10 MiB. With malloc trimming off, the unified open is 0.98× the pre-M0a one.
- The DOCX eager full-text row also carries S8a's paragraph separator (B3, an accepted change), which costs +2.77% instructions on the first call.
- Budgets: in a fresh process the unified XLSX open costs +8.9 µs over the leaf calls and the DOCX source open +10.0 µs, both within 20 µs. DOCX `list_paragraphs` costs 1.3× the leaf for the `compat` wrappers; M1's node view replaces them.
- Memory (R5):
  - DOC (path and bytes) and PPT (bytes) hold +1.01× to +1.94× the input size after open, which is the input plus its CFB index.
  - Peak memory during open is unchanged, and `retain_source(false)` removes the retention.
  - Other formats hold +0.6 to +4.3 KB more.
- Harness fix: the DOCX leaf full-text leg now pairs with `File::text`'s separator (`d10dd4c747`), and an oracle test pins it (`c475a527df`).
- Owner decisions (record §9):
  - F1: accept E1, or give fresh-child routes a warm-up call, or gate first calls with standalone binaries.
  - F2: accept E2, or fix the malloc tunables, or add bare-loop variants.
  - S10 recommends accepting both.
- Follow-ups:
  - F3: drop S1's whole-child RSS tripwire; it measures the `git` child the harness runs, not the harness.
  - F6: keep future captures off the shared tmpfs.
- Merged into `uo/m0a` (`80f2642b01`). Coordinator fix (`31d6b31d2f`): the oracle test wrote under `std::env::temp_dir()`, which the source policy forbids in the unified-ops modules. It now writes under `filesystem::scratch_root`.

## Landing
- Base (DOCX section commit-target reuse, perf records) merged into trunk (`08c26f7f3d`), then trunk into `uo/m0a` (`4e16969ffd`), with no conflicts.
- Gates on the merged `uo/m0a`:
  - clippy: litchi-docx, ops, the facade, litchi-py and the harness;
  - tests: litchi-docx (2040), litchi-ops (445), the facade targets, Python and the harness (633);
  - rustdoc (litchi, ops; all features);
  - the boundary and non-iWork gates and their tests, and the perf policy and CRUD index tests.
- Workspace fmt fails only on the base's iwa-archive and soapberry-zip, as before.
- The OfficeCLI delta table is an M1 tool and is not on `uo/m0a`.
