# 0248 — Coordinator restart plan (survey of trunk 97ba48dcd7)

Status: planning record, proposed. Surveyed on 2026-10-13 from the detached, read-only worktree `/home/zhuhe/code/litchi-worktrees/uo-survey`.

Trunk `feat/unified-ops-api` is 356 commits ahead of base `feat/office-format-completeness` (32ad64e320) and 0 behind it, so no base merge is pending. Records 0016–0247 are all present in the tree. The "commit blocked" notes in PROGRESS.md (for example rows 0053–0127) describe how they were written at the time; they are not missing commits.

Paths are relative to the worktree, and `ops/` means `crates/litchi-ops/`.

Legend:
- **S**: supported.
- **P**: partial (what is missing follows).
- **A**: absent.
- "pending()": the verb falls through to the `_ => Err(pending(..))` arm, which returns `unsupported_capability` (`ops/src/write.rs:94,149,211,267,346,439`).

---

## 1. Format × operation matrix (verified from code and tests)

The dispatch arms are the authority:
- `ops/src/write.rs`: set covers 8 formats, add 5, remove 7, move 6, swap 6, publish 8.
- `ops/src/read.rs:68-121,202-237,262-371`: get/find/query/view cover 8 formats.
- `ops/src/save.rs:269-315` and `ops/src/create.rs:168-245`: 8 formats each.
- `ops/src/transaction.rs:475-921`: the edit closures. `batch/execute.rs:71,123` inherits the transaction coverage.

The `File` enum has 11 format variants: Docx, Xlsx, Pptx, Xlsb, Doc, Xls, Ppt, Rtf, Odt, Ods, Odp (`ops/src/file.rs:128-170`).

| Fmt | open | get/find/query/view | caps | set | add | remove | move | swap | batch | save | create |
|---|---|---|---|---|---|---|---|---|---|---|---|
| docx | S | S [1] | S | P [2] | P para/run | P para/run | P | P | S | S | P [3] |
| xlsx | S | S | S | P [4] | P sheet | P worksheet | P sheet order | P sheet order | S | S | P [3] |
| pptx | S | S | S | P [5] | P slide, text-box | P [6] | P slides | P slides | S | S | P [3] |
| xlsb | S | S | S | P [7] | P sheet | P sheet, cell | P sheet order | P sheet order | S | S | P [3] |
| doc | S | P [8] | P [9] | P para `text` only | A | A | A | A | P set only | S | P blank |
| xls | S | P [10] | S | P [11] | A | P cells/ranges | A | A | P | S [12] | S |
| ppt | S | P [13] | S | P [14] | A [15] | P slides | P slides | P slides | P | S [12] | S |
| rtf | S | P [16] | S | P [17] | P para (text), run | P para, run | P main-story para | P | S | S | S |
| odt | P [18] | A [19] | P [20] | A | A | A | A | A | A | A | A |
| ods | P [18] | A [19] | P [20] | A | A | A | A | A | A | A | A |
| odp | P [18] | A [19] | P [20] | A | A | A | A | A | A | A | A |
| odg, odc, odi, odb, oth, odm, odf | A [21] | A | P [22] | A | A | A | A | A | A | A | A |
| pages, keynote, numbers | A [23] | A [23] | P [22] | A | A | A | A | A | A | A | A |

Totals:
- **Read through get/query/view:** 8/21 formats.
- **Any write:** 8/21 formats.
- **Whole-format `unsupported_*` still returned:** 13/21 (3 ODF read-only, 7 ODF, 3 iWork). Every one of these violates the owner override (`docs/UNIFIED_OPS_GOAL.md:36-44`).

Notes:
1. **docx reads.** Readable elements are the document (10 core props), paragraph (28 props), run (38), hyperlink, table/row/cell, sdt, equation and section (`ops/src/format/docx/spec.rs:468-967`).
   - Read gaps: header, footer, foot/endnote, comment, style, numbering, settings, theme, bookmark, field, picture, chart, ole, toc, revision, textbox and shape (`spec.rs:759-789`).
   - All formats render the text, annotated, outline and stats views. The issues, forms, html, svg and pdf views are gaps (`ops/src/capabilities.rs:2305-2337`).
2. **docx writes.** set, add and remove only reach direct `/body` paragraphs and runs. Paragraphs inside cells, SDTs or hyperlinks are refused (`ops/src/format/docx/write.rs:613-631`, `node.rs:404-462`, `structure.rs:107-325,809-819`).
   - Tables, rows and cells are read-only.
   - Tests: `ops/tests/docx_*.rs` (about 116 cases).
3. **create.** Non-core metadata, templates and Strict conformance are refused (`ops/src/format/docx/create.rs:33-63`, `capabilities.rs:1447-1488,2996-3004`, `format/xlsx/create.rs:27`).
4. **xlsx set.** Cells accept 51 setters: value, formula, clear, font, alignment, protection, numFmt, fill and borders. Ranges add merge. Sheets accept name and hidden; row and column accept height/width, hidden and outline (`ops/src/format/xlsx/spec.rs:226-278,836-968`).
   - Gaps: comment, chart, pivot, data validation, conditional formatting, autofilter, hyperlink, picture, shape, freeze panes, zoom, tabColor, print area and computedValue (`format/workbook/spec.rs:481-562,900-1003`).
5. **pptx set.** Shapes accept only text, x, y, width and height; slides accept name and hidden (`ops/src/format/pptx/spec.rs:1445-1549`).
   - Rich paragraph/run formatting, table cells, pictures, notes and comment writes are absent (`capabilities.rs:3095-3104`).
6. **pptx remove.** Removes top-level slides and shapes only. Slides with dependencies, the last slide and grouped children are refused (`capabilities.rs:2773-2789,3045-3094`). Decks with p14 sections are refused by ordering (`crates/litchi-pptx/src/presentation/order.rs:1397-1400`).
7. **xlsb set.** Cells and ranges accept only value and clear; sheets accept name, hidden and visibility. Formula, formatting (about 40 props) and merge are gaps (`ops/src/format/xlsb/spec.rs:21-80,336,433-800`).
8. **doc reads.** Only the document, body and `/body/p[N]` with `text` are readable (`ops/src/format/doc/spec.rs:11-106`). Runs, tables, sections, headers and styles are absent.
9. **doc capabilities.** `SPEC.gaps` is empty (`ops/src/format/doc/spec.rs:105`), so the missing DOC elements are not enumerated in `capabilities()`, as the override requires.
10. **xls reads.** Workbook, sheet, cell, range, row and column are readable. Table, named range, comment, chart, conditional formatting, hyperlink, merge and formula are gaps (`ops/src/format/xls/spec.rs:218-350`).
11. **xls set.** Accepts cell value and clear, and sheet name. Formula cells are refused (`format/xls.rs:113-125`, `xls/write.rs:203-208`).
    - The leaf already has `insert_cell` and `rename_sheet` (`crates/litchi-xls/src/cell_values/mod.rs:2573,3220`).
12. **Unlocked encrypted XLS/PPT.** These refuse plaintext publication (`capabilities.rs:637-660`).
13. **ppt reads.** Presentation, slide, shape, placeholder, picture, table, group and connector are readable. Notes, paragraph, run, cell, chart, master and layout are gaps (`ops/src/format/ppt/spec.rs:22-57,346-481`).
14. **ppt set.** Shape text and geometry, and slide hidden (`format/ppt/spec.rs:913-925,1012`).
15. **ppt add.** "PPT insertion remains a write gap" (`ops/tests/capabilities.rs:733-739`).
16. **rtf reads.** Document, body, paragraph and run are readable. Table, row, cell, header, footer, note, comment, picture, hyperlink, style and numbering are gaps (`ops/src/format/rtf/spec.rs:14-23,399-438`).
17. **rtf set.** Paragraphs accept text, align, spacing, indents, keep options and pageBreakBefore. Runs accept 13 character properties (`format/rtf/spec.rs:134-355`). Records 0232–0247.
18. **ODT/ODS/ODP open.** Zip packages open; flat (single-XML) ODF is refused (`ops/src/open.rs:547-572`, `open/bytes.rs:39-41`). Only handle reads work: text, counts and metadata (`format/odt.rs:86-128`, `ods.rs:110-194`, `odp.rs:83-128`).
19. **ODT/ODS/ODP get/query.** These are refused with "M1 R7 adapter has not landed" (`capabilities.rs:1003-1014`).
20. **ODT/ODS/ODP capabilities.** The report is a whole-format refusal.
21. **Other ODF open.** Detection recognises these formats, then `not_enabled` returns `unsupported_format` (`ops/src/open.rs:331-346`, `open/path.rs:490-493`).
22. **Other ODF and iWork capabilities.** `closed("M6", ...)` returns `UnsupportedFormat` (`capabilities.rs:1030-1037`).
23. **iWork.** Open only through the facade's read-only `litchi::iwork::Document` (`crates/litchi/src/iwork/model.rs:19-169`). The unified open refers callers there with `iwork_unsupported` (`ops/src/error.rs:785`).
    - `litchi-ops` must stay free of iWork dependencies (`ops/Cargo.toml:22-44`, `tools/non_iwork_gate.py`).

### P0 surface (design §13, `docs/UNIFIED_OPS_API_DESIGN.md:276-288`) — missing items

| P0 item | docx | xlsx | pptx |
|---|---|---|---|
| Full basic CRUD | P: paragraph/run only; tables, cells and nested stories read-only or absent | P: cell, range, row, col, sheet | P: slide, shape text and geometry |
| Worksheet lifecycle | n/a | **present** (0102); veryHidden not settable | n/a |
| Metadata core/app/custom | core read-only, set only at create; app and custom absent | read-only; refused at create | read-only; refused at create |
| Style read + apply | apply `style`/`rStyle` works; reading styles is a gap (`spec.rs:768`) | cell formatting only | absent |
| Numbering/list basics | `numId`/`numLevel` set; numbering parts are gaps | n/a | `list` is a gap |
| CSV/TSV import | absent (`batch/execute.rs:322`: `import` pending) | absent | n/a |
| Formula write + eval | n/a | write present; eval absent (no `litchi-eval` dependency; `computedValue` gap) | n/a |
| Comments/notes | absent | absent | read only |
| Hyperlinks | read only | absent | absent |
| Pictures | absent | absent | read, plus remove as shape |
| find/replace, merge, refresh (new engines) | absent everywhere (no matches for replace, merge or refresh in `ops/src`) | | |
| Chart/SmartArt CRUD (v1) | absent | absent | read and remove frame only |

Two P1 items have already landed early: XLSX merge/unmerge (0104) and row/column outline (0103).

## 2. ADR 0034 rollout status

**W0 gate: done** (0017). `tools/check_api_glossary.py` builds rustdoc JSON per profile and does two things:
- It bans exact names under a path scope. With `recursive:false` it checks only direct children (`:291-299`).
- It requires each registered family module to export all six of `Edit`, `Patch`, `Commit`, `Diagnostics`, `Error` and `LimitKind` (`:28,300-307`).

It validates its policy strictly (`:37-98`), and CI runs it (`.github/workflows/rust-ci.yml:282-287`).

There are 18 profiles in `tools/api_glossary.json`. This run reported "18 profiles, 574 banned names, 5 transaction families checked", exit 0.

| Profiles | Bans | Families |
|---|---|---|
| core-minimal | 0 | — |
| docx-minimal, docx-full | 40 each | none |
| opc-minimal, opc-full | 23 each | none |
| pptx-minimal | 79 | none |
| pptx-full | 70 | none |
| xlsx-minimal, xlsx-full | 8 each | none |
| xlsb-minimal, xlsb-full | 32 each | none |
| drawingml | 5 | none |
| doc-minimal, doc-full | 34 each | `doc::package::edit` |
| xls-minimal | 24 | `xls::package::edit` |
| ppt-minimal | 35 | `ppt::package::edit` |
| rtf-minimal | 86 | `rtf::tail_append` |
| ops-rtf | 1 | — |

The checker runs on stable 1.95 through `RUSTC_BOOTSTRAP=1` (`check_api_glossary.py:351`), so **no nightly toolchain is needed**. It does depend on the unstable rustdoc JSON schema, so the toolchain pin matters.

Holes in the gate:
- **pptx-full is missing 9 primary-transaction bans** that pptx-minimal has: `opened::Transaction`, `opened_presentation_transaction`, `apply_opened_presentation_{commit,patch}`, and others.
- **No family is registered for docx/pptx/xlsx/xlsb.** docx appears to export all six types. pptx `opened` lacks `Error` and `LimitKind`. xlsb `cell_values` lacks `LimitKind`. xlsx `workbook::edit` has only Edit, Commit and Patch.
- **No full-feature profile** for xls, ppt or rtf.
- **No profile at all** for the facade `litchi`, `litchi-ooxml-common`, ODF, `litchi-sheet`/`word`/`slide`, `ole-common`, `cfb`, `biff` or `crypto`. `litchi-ops` is profiled only for one RTF name.

Per format group:

| Wave / format | owner §1 | txn §2 | open/create/save §3 | limits §4 | errors §5 | metadata §6 | status |
|---|---|---|---|---|---|---|---|
| W2 docx (0018–0046) | P (`document_mut` at `litchi-docx/src/package/package/document.rs:480`) | primary done, family unregistered | done (0038–0044) | P | P | pending (`props`/`props_mut`/`put_props`/`clear_props`, `package/model.rs:937-986`) | **partial** |
| W2 pptx (0116–0164) | P | P | P (`open_with_limits` at `package/codec.rs:211`; `from_vec`) | pending | pending (`StaleSource`, about 107 refs) | pending | **partial** |
| W2 xlsx | pending (`Workbook::save/edit/apply`, `workbook/model.rs:1349-1509`) | pending | pending (`Package::create`, `from_slice`, `from_bytes(Vec)`, `open_with_limits`, `into_workbook` in `package.rs:39-309`) | SourceLimits only | pending | pending | **mostly pending; the deepest reshape** |
| W2 xlsb (0165–0187) | P | P | P (`edit_with_limits`, `open_with_limits`, `from_slice`, `into_workbook`, `raw::Writer` in `package/mod.rs:132-633`) | P (nested `ReadLimits`) | pending | pending | **partial** |
| W3 ODF | pending (ods `Spreadsheet`, odt `Document`) | pending | pending; save is non-atomic `fs::write` (`litchi-odt/src/builder/package.rs:71`) | pending (`SourcePackageLimits` with `ReadLimits` aliases) | pending | — | **not started** (no records) |
| W4 doc (0188–0193) | done | done | P (part-level `open_with_limits` ×8) | done | limits done | pending | **mostly done** |
| W4 xls (0194–0209) | done | done | done | done | limits done | pending | **mostly done** |
| W4 ppt (0210–0221) | done | done | done | done | limits done | pending | **mostly done** |
| rtf (0222–0247, no wave) | done | done | done | done | done | — | **done** |
| W1 iWork | pending; owners must make the change (`docs/unified-ops/IWORK_API_REQUESTS.md`); `Document` is still in numbers and keynote | | | | | | **needs owner** |

Remaining banned or legacy public names. These come from a source grep; the gate's own results are authoritative.
- **xlsx:** `Package::create`, `from_slice`, `into_workbook`; `WorkbookPatch`/`WorkbookCommit` in pivot modules.
- **xlsb:** `edit_with_limits`, `from_bytes_with_limits`, `raw::Writer`.
- **pptx:** `open_with_limits`, `from_vec`, `rollback` ×3, `StaleSource`.
- **docx:** `document_mut`, `props*`, `write_to_stream`.
- **opc:** `open_with_limits`, `to_stream`, `write_to_stream`, `from_vec`.
- **ooxml-common:** `Props` (`properties.rs:156`).
- **odf-common/odt/ods/odp:** `SourcePackageLimits`, `open_with_limits`, `from_vec`, ods `apply_document_patch`/`edit_document`/`document_mut`.
- **ops:** `open_with` (`ops/src/open.rs:90`), `edit_with_limits` (`transaction.rs:284`).
- **core:** `Limits` (`budget.rs:69`; core-minimal registers no bans).
- **Scoped `Transaction` types remaining:** docx 15, xlsx 20, pptx 12, xlsb 13, doc 11, xls 11, ppt 12, ods 10.

## 3. Other surfaces

**Python (`crates/pyo3-litchi`).**
- What exists: the legacy read classes (`Document`, `Presentation`, `Workbook`, ...) built over `litchi::open` (`src/document.rs:64`).
- Done:
  - The `LitchiError` hierarchy with code, message and suggestion, plus subclasses (`src/errors.rs:1-40`).
  - The GIL is released on reads (`py.detach`).
  - A `.pyi` file with a stub-sync test.
  - 132 tests.
- **Missing (all of M3):**
  - `litchi.open`/`create`/`open_bytes` returning a unified handle.
  - get, query, view, set, add, remove, move, swap, batch, apply, dump, save, close and capabilities.
  - Context manager.
  - Dotted-key dict props.
  - Node-view dict returns.
  - Zero-copy `raw()` bytes.
  - A pytest mirror of the Rust scenarios.
- Branch `uo/m1` 668d66883c holds the R9 read surface (§5).

**CLI.** **Absent.** There is no `[[bin]]` outside `tools/`, and no binary uses clap. The goal requires one (`docs/UNIFIED_OPS_GOAL.md:202-205`, M3) and so does design §11.

**OfficeCLI parity.**
- `tools/officecli_capability_delta.py` generates `docs/unified-ops/officecli-capability-delta.md` (4,457 lines) from the 152 OfficeCLI schemas and `ops/tests/golden/capabilities/*.json`. It has a `--check` mode and 14 unit tests.
- It reads the capability goldens, **not the registry**. Its header still says "M1 is read-only", so write support is not shown.
- There is **no harness** for the 96 OfficeCLI `examples/*.sh` scripts; it is only planned (record 0015 W8, `0015-m2-write-kernel-plan.md:1646`).

**Paired benchmarks** (`tools/perf-baseline`: `src/unified_ops.rs` plus `unified-*` write packages; 91 result files in `docs/unified-ops/results/`).

Budgets met:
- 1k batch within 2×: docx (0078/0091), xlsx (0101), pptx (0156/0158), xlsb (0172/0180), xls (0207), ppt (0221), rtf (0247).
- Single set within 10%: pptx (0155/0163), xls (0205), xlsb (0174).

Budgets missed:
- Single set: docx (0089), doc (0188, up to 20.4%), ppt (0216, 6–39%), rtf on tiny inputs (0236/0239/0246).
- Read overhead: doc (0190, 13.6%), xls (0203).
- 10k linear scaling: ppt (0218/0220).
- docx 10k batch is a typed refusal (0089).
- RTF open (E2) and the M0a F1/F2 explanations are awaiting the owner.

**Not measured at all:**
- Paired save with zero added copies.
- Query early termination.
- XLSX single-set ratio.
- Anything for ODF or iWork writes.

**Fuzz.**
- Selector fuzzing is done: `crates/litchi-core/fuzz/fuzz_targets/selector_grammar.rs` checks the parse → display → parse round trip.
- `ops/fuzz` covers paths, values and query_plan.
- **Missing:** a batch-payload (OfficeCLI JSON) ingress fuzz target.

**Goldens and schema export.**
- Capability goldens cover 9 profiles, plus node-view goldens.
- **Missing:** a property-registry snapshot and generated JSON schema export (M3).

**Facade.**
- `litchi::open`/`open_bytes`/`create` come from `crates/litchi/src/unified.rs:62-121`.
- The old wrappers were deleted in S9 (`docs/unified-ops/s9-deletion.md`).
- `litchi::iwork::Document` is intentionally retained.
- `IWORK_API_REQUESTS.md:10` cites a deleted path (`crates/litchi/src/document/types.rs`).

## 4. Health check of trunk 97ba48dcd7 (2026-10-13)

Environment:
`CARGO_TARGET_DIR=<wt>/target CARGO_INCREMENTAL=0 CARGO_PROFILE_{DEV,TEST}_DEBUG=0 TMPDIR=/tmp/claude-1001/uo-t/survey`, with `3rdparty` symlinked into the worktree (gitignored). Toolchain 1.95.0 (pinned).

| Check | Exit | Result |
|---|---|---|
| `cargo fmt --all --check` | **1** | 205 hunks in `crates/litchi-rtf/src/{edit.rs (21), edit/composition.rs (73), edit/transfer.rs (9), tail_append.rs (102)}`. All are missing match-arm trailing commas (`match_block_trailing_comma = true`). They come from the 2026-10-12 RTF commits 3327f64411, bcf6330558 and d13ab39512. |
| `cargo clippy -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | **101** | `error[E0432]: unresolved import litchi_ppt::writer::Writer` at `crates/litchi/examples/odraw_native_smoke.rs:11`. Record 0213 (d45c1f3f77) replaced the PPT `Writer` with `Builder` but did not migrate this example (`crates/litchi/Cargo.toml:305-307`, required features doc+ppt+xls). With `--keep-going` this is the only error. The same command with `--lib --tests --benches` exits **0**. |
| `cargo test -p litchi-ops --all-features` | **0** | 1,764 passed, 0 failed, 0 ignored (76 suites incl. doctests). |
| `cargo test -p litchi --all-features` | **101** | The same example compile error stops the build. `--lib --tests` exits **0** with 240 passed; `--doc` exits **0** with 20 passed and 1 ignored. |
| `python3 tools/check_crate_boundaries.py` | **0** | 66 packages, 268 dependency declarations, 11 explicit debt items (iWork). |
| `python3 tools/non_iwork_gate.py verify` | **0** | Verified 47 bulk roots, 35 facade-safe trees and the combined tree. |
| `python3 tools/check_api_glossary.py --target-dir <wt>/target/glossary` | **0** | 18 profiles, 574 bans, 5 families. Stable plus `RUSTC_BOOTSTRAP`; no nightly needed. |
| `crates/pyo3-litchi/tests/run_tests.sh` | **0** | 132 tests OK (Python 3.14.4). |

Consequence: CI `rust-ci.yml` is red on trunk, because of fmt (`:290`) and `cargo check --workspace --all-targets` (`:301`). Clearing it is step 1 below; it is small.

## 5. Stale-branch inventory

The `wip(reference)` commits on the four reference branches say "stopped before final gates, not verified". Trunk redid M1/M2 as large commits (723f8c13eb, 0314261beb, caf1d800c5, 724a37462f, 8fe6801d8d), so `git cherry` matches almost nothing. Every port has to adopt ADR 0034's names (`Workbook`→`Package`, `Key`/`Placement`→`Selector`/`Position`).

| Item | Trunk equivalent | Verdict | Size |
|---|---|---|---|
| `uo/m2-w3-executor` 5258f8b42d (generic `Handler` executor, bisection, `BatchReport`) | yes: `ops/src/batch/execute.rs:43-209` (per-format coalescers, `BatchMode`) | **obsolete**. Adapt only its property test (a coalesced batch equals the same items applied one at a time, byte-for-byte). | S (~400 LOC tests) |
| `uo/m2-l-pptx-w4` 9474e2b6ce (add/remove/move slide) | partial: insert, move, swap and remove are on trunk (0136–0150, `litchi-pptx/src/opened/transaction.rs:129,163`) | **port two pieces**: (a) layout placeholders created in a new slide; (b) p14 section maintenance (trunk refuses section decks, `presentation/order.rs:1397`) | M (~1.5k src + 1k tests) |
| `uo/m2-l-xlsb-w1` 876588a23c (`&self` fork_* forms) | yes: package transactions (0165–0187) | **obsolete**, except the merged-range forks (~100 LOC), which could close xlsb `MERGE_GAP` | S |
| `uo/m1-r4-perf` 012fd509ff (lazy comparisons, `Shortcut` plans, query-corpus differential test) | none (`Shortcut`, `query_corpus` absent; trunk drifted +863/-287 in query/value) | **port**, with the differential corpus test first | M (~1.5k src + 0.7k tests) |
| `uo/m2` b31129653f (`litchi_ole_common::property_set::SummaryProperties`, DOC/PPT/XLS writers) | none; create refuses initial metadata (`ops/src/format/doc.rs:82-85`, `ppt.rs:130`, `xls.rs:160`) | **port**. ole-common applies cleanly (~1.3k); writer wiring conflicts with Builder/immutable owners. | M (2–3 days) |
| `uo/m1` 668d66883c (R9 Python read surface: `file.rs`, `nodes.rs`, `pyser.rs`, +418 `.pyi`) | none (no get/query/view in `crates/pyo3-litchi/src`) | **port**, adapting to trunk's `Node` API | M (~1.5k src + 1.2k tests) |
| `uo/m2` 4dba739fac (OOXML document properties, DOCX list definitions, +8.7k) | none (`litchi-ooxml-common/src/document_properties/` absent) | **port**: P0 metadata and numbering | L (4–6k after dedup) |
| `uo/m1` legacy readers (DOC stories/headers/comments/fonts/Word97 tables; XLS defects; PPT outline/notes; 0x93/0x94 quotes, `litchi-doc/src/parts/text.rs:547`) | none | **port** for tier 2 reads; the quotes fix is a 20-line S | L (~8k total) |
| `uo/m2` W1/W2/W4 ops kernel, DOCX/XLSX/PPTX W1–W7 | trunk has its own versions | **obsolete** | — |

## 6. Prioritized plan to the Definition of Done

Sizes are agent-days: S ≤1, M 2–4, L 5–10, XL >10.

"Owns" lists the files each stream may touch; streams marked ∥ touch disjoint files and can run at the same time. Shared files (`ops/src/{write,read, transaction,save,create,capabilities}.rs` dispatch arms, `ops/src/file.rs`, `ops/Cargo.toml`, PROGRESS.md) change only in short, coordinator-merged hunks. The tier order is OOXML → OLE → RTF/ODF → iWork.

### Phase A — trunk green and gate hardening (blocking, serial, ~2 days)

1. **Fix the red CI.** (S)
   - Run `cargo fmt` on `crates/litchi-rtf`.
   - Migrate `crates/litchi/examples/odraw_native_smoke.rs` to `litchi_ppt::Builder`.
   - Re-run the §4 suite plus `cargo check --workspace --all-targets --all-features`.
   - Owns: `litchi-rtf` (format only), `crates/litchi/examples`.
2. **Harden the glossary gate.** (S) Depends on 1.
   - Add the 9 missing bans to pptx-full.
   - Register the docx `document` family.
   - Add full-feature profiles for xls, ppt and rtf.
   - Add core bans for `Limits`.
   - Add the facade and ooxml-common profiles where rustdoc globs allow (0017 caveat).
   - Owns: `tools/api_glossary.json`, `docs/API_GLOSSARY.md`.
3. **Make capabilities truthful.** (S) Can run alongside 2.
   - Enumerate the DOC element gaps in `ops/src/format/doc/spec.rs:105`.
   - Regenerate the delta table with write-aware statuses and fix its "M1 read-only" header.
   - Fix the stale `IWORK_API_REQUESTS.md:10` cite.
   - Owns: `ops/src/format/doc/spec.rs`, `tools/officecli_capability_delta.py`, `docs/unified-ops/officecli-capability-delta.md`, golden capabilities.

### Phase B — Tier 1 OOXML to P0 (parallel; ADR 0034 W2 before feature work in each crate)

4. **xlsx W2 reshape.** (L)
   - Demote `Workbook` to a projection.
   - Make `Package` the owner of edit/apply/save.
   - Remove `Package::create`, `from_slice`, `from_bytes(Vec)`, `into_workbook` and `WorkbookPatch`/`WorkbookCommit`.
   - Register the family.
   - Adapt `ops/src/format/xlsx`. This must land before 7b.
   - Owns: `crates/litchi-xlsx`, `ops/src/format/xlsx`.
5. **Finish docx/pptx/xlsb W2 residue.** (M each) ∥ by crate.
   - docx: `props*` → metadata Edit; `document_mut`; `write_to_stream`.
   - pptx: `open_with_limits`, `from_vec`, `StaleSource`, family `Error`/`LimitKind`.
   - xlsb: `*_with_limits`, `raw::Writer`, nested `ReadLimits`, `LimitKind`.
   - opc: the stream verbs.
6. **Shared OOXML metadata (P0).** (L) Port 4dba739fac's core/app/custom properties onto `litchi-ooxml-common/src/properties`, then expose the `/metadata` set path and create-time metadata.
   - Owns: `litchi-ooxml-common`, plus one `metadata` module per ops format.
   - Each format stream picks it up afterwards.
7. **Per-format P0 streams** ∥ (each owns its leaf crate and `ops/src/format/<fmt>/`):
   - **7a docx (XL).**
     - Paragraph/run writes inside cells, SDTs and hyperlinks.
     - Table, row and cell CRUD.
     - Hyperlink write.
     - Comments.
     - Headers and footers (P1, but they unlock stories).
     - Picture add.
     - Style read.
     - Numbering parts (from 4dba739fac's `numbering/lists.rs`).
     - The 10k batch refusal (0089).
     - Single-set overhead (0089 miss).
   - **7b xlsx (L), after 4.**
     - Comments.
     - Hyperlinks.
     - Pictures.
     - Freeze panes.
     - Named-range write.
     - veryHidden.
     - Single-set paired benchmark.
   - **7c pptx (L).**
     - Rich paragraph/run set.
     - Table cell text/fill.
     - Picture add/set.
     - Notes and comment writes.
     - Port 9474e2b6ce placeholders and sections.
   - **7d xlsb (M).**
     - Cell formatting set.
     - Formula write.
     - Merges (fork logic).
     - Named ranges.
8. **Query performance port** (012fd509ff) and a paired query early-termination benchmark. (M) ∥
   - Owns: `ops/src/query/`, `ops/src/value/`, `ops/src/node/path.rs`.
   - Coordinate the xlsx/pptx `node.rs` hunks with 7b/7c.
9. **Benchmark gaps.** (M) ∥
   - Paired save scenario (zero-copy claim).
   - XLSX single-set ratio.
   - Batch property test from 5258f8b42d.
   - Batch-ingress fuzz target.
   - Owns: `tools/perf-baseline/unified-*` (new route names only), `ops/fuzz`, `ops/tests/batch_*`.

### Phase C — Tier 2 OLE read+write loop (starts after Phase A; ∥ with B; disjoint crates)

10. **Legacy metadata.** (M) Port the `SummaryProperties` merge (b31129653f) into `litchi-ole-common`, wire it into the doc/xls/ppt Builders, and lift the create-metadata refusals.
11. **doc (L).**
    - Port the `uo/m1` DOC readers: stories, headers, comments, fonts, Word97 tables and the quotes fix.
    - Run reads and run/character set, wrapping the existing leaf edits (`litchi-doc/src/package/edit.rs:540-594`).
    - Paragraph add/remove/move/swap.
    - Read and single-set overhead (0188/0190 misses).
12. **xls (M).**
    - Sheet add/move/swap/remove.
    - Rename via the existing leaf API.
    - Row/col.
    - Formula cells.
    - Read overhead (0203 miss).
13. **ppt (M).**
    - Slide add (leaf `add_slide` exists in the writer model).
    - Notes and paragraph reads.
    - 10k scaling (0218/0220) and single-set misses (0216).
    - Port the `uo/m1` PPT outline/notes readers.

    W4 residue for 11–13: part-level `open_with_limits`, `rollback`, metadata §6.

### Phase D — M3 consumption surfaces (∥ from now on; consumes the ops API only)

14. **Python.** (L) Port R9 (668d66883c). Then add:
    - the unified `File` handle and `open`/`open_bytes`/`create`;
    - the write verbs, `apply`/`batch`/`dump`;
    - the context manager and dict props;
    - GIL release around open, query, batch and save;
    - zero-copy `raw()`;
    - the pytest mirror and `.pyi` sync.

    Owns: `crates/pyo3-litchi`.
15. **CLI crate** (e.g. `crates/litchi-cli`). (M)
    - Same verbs, `--json` equal to the node view.
    - Register it in the ADR 0024 inventory, `tools/crate_boundaries.json` and `non_iwork_gate.py`.
16. **Schema export and registry golden snapshot.** (M) The delta tool should read the registry export.
    - Owns: `ops/src/registry`, `ops/tests/golden`, the tool.
17. **OfficeCLI example-corpus harness** (record 0015 W8) over the CLI. (L) Depends on 15. The v1-scope subset goes green in waves, with the deltas recorded.

### Phase E — M4 engines (after B reaches P0 CRUD)

18. **CSV/TSV import** (M) and **formula evaluation via `litchi-eval`** (M; a new ops dependency, which needs a boundary update).
19. **Chart and SmartArt typed CRUD.** (L)
20. **Three new engines** (XL, split M4a/M4b):
    - formatting-aware find/replace (`find=`/`replace=` on `set`);
    - template `merge`;
    - TOC/field `refresh`.

### Phase F — Tier 3 (starts after Phases B and C have a working loop)

21. **RTF to parity** (L).
    - Tables, headers, notes, hyperlinks, pictures and styles.
    - Tiny-input overhead (0236–0246).
    - RTF already has a full paragraph/run loop.
22. **W3 ODF leaf wave** (L).
    - Package split.
    - Atomic save (replace `fs::write`).
    - `SourcePackageLimits`→`SourceLimits`.
    - ods `Spreadsheet` and `edit_document`.
    - Owns: `litchi-odf-common`, `litchi-odt`, `litchi-ods`, `litchi-odp`, `litchi-od*`.
23. **ODF adapters.** (L each; ∥ by format after 22)
    - First odt, ods and odp (M1 R7 read plus write): the leaf `Edit`/`commit` and `Builder` APIs already exist.
    - Then add `File` variants for odg, odc, odi, odb, oth, odm and odf, which have edit/commit/builder APIs (record 0002:88-100).
    - Flat ODF inputs.
24. **iWork, M6** (XL). Needs **decisions D1 and D2** below.
    - Leaf requests are still open: creation, sheet/table lifecycle, slide add, shape CRUD, Pages sections (`IWORK_API_REQUESTS.md:38-88`).

### Phase G — M5 acceptance closure

25. CRUD checklist evidence per format.
26. Full-corpus open → empty edit → save byte regression.
27. The benchmark matrix with every miss explained.
28. Delete `litchi::iwork::Document` once iWork joins `File`.

### Critical path and parallelism

The critical path is 1 → 4 → 7b → 18 → 20 → 25–27. Tier 3 adds 22 → 23 → 24.

After step 1, these streams can run at once:
- 5 (per crate)
- 6
- 7a and 7c
- 8 and 9
- 10–13 (OLE crates)
- 14, 15 and 16

4 must precede 7b. 6 must precede the metadata parts of 7a–7d and 10. 15 must precede 17.

Phase F starts only once the tier 1 and tier 2 loops are signed off (see D3).

### Decisions that need the owner

- **D1 — iWork placement.** *Owner decision 2026-10-15: remove the iWork-free gate; `File` gains the iWork variants directly (no facade wrapper, no adapter crate).* `litchi-ops` must stay iWork-free (`non_iwork_gate`, constraint 7). The "single `File` enum for all 21 formats" therefore needs one of:
  - (a) a facade-level `File` that wraps `litchi_ops::File` plus iWork variants;
  - (b) relaxing the gate for a dedicated `litchi-ops-iwork` crate;
  - (c) a trait-free adapter in the facade.

  Recommended: (a).
- **D2 — W1 iWork renames and the API requests.** *Owner decision 2026-10-15: the iWork owner agent is paused; coordinator agents may modify the iWork crates.* These need scheduling with the iWork owners. Without leaf `Package::new` and lifecycle APIs, iWork can meet "write" only at element granularity.
- **D3 — Tier gating.** RTF tier 3 work (0222–0247) ran while DOC still had paragraph-text writes only. Confirm whether the tier 1/2 "working read+write loop" is met now (all 7 have set/save/create), so that Phase F may proceed in parallel, or whether it should wait for P0.
- **D4 — Outstanding provisional decisions** (PROGRESS.md:244-268):
  - M0a S10 F1/F2;
  - record 0015 D1–D19 / Q1–Q14;
  - record 0010 Q3 (adding the `regex` crate);
  - Q6 (case-insensitive sheet names, a behaviour change for XLSB/ODS);
  - record 0007 defaults.

  All are still awaiting confirmation.
- **D5 — Budget misses.** Accept or explain them: doc read 13.6% (0190), ppt single-set up to 39% (0216), ppt 10k scaling, and RTF tiny-input overhead. These need either engineering (planned above) or an owner-approved, recorded explanation.

### Not claimed

This record:
- changes no code;
- makes no performance claim beyond citing records;
- verifies the matrix from dispatch, spec and test code, not by running each verb;
- leaves size estimates as planning guesses.
