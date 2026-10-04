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
