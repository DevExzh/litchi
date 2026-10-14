# 0279 — Field and table-of-contents refresh (plan step 20c, M4 engine 3 of 3)

Status: implemented on `uo2/s20-refresh` (cut from `feat/unified-ops-api` at
80e24f8745); awaiting coordinator review and merge.

Plan: [0248](0248-coordinator-restart-plan.md) Phase E step 20. Design: §13
P0 "`refresh` (TOC/field recalculation — the docx TOC writer today emits an
'update this field' placeholder, with no entry computation)" and the closed
error list's `refresh_failed`. Glossary: [API_GLOSSARY.md](../API_GLOSSARY.md),
section "Field and table-of-contents refresh (record 0279)". The M4 shape
follows [0274](0274-template-merge.md) (one mutation, at most one native
commit, `transaction::Commit<Summary>`).

## 1. What landed

### API (`litchi_ops::refresh`, re-exported as `litchi::refresh`)

- `File::refresh(&refresh::Options)` and `File::refresh_with_limits(..,
  &SemanticLimits)` return `transaction::Commit<refresh::Summary>`;
  `transaction::Edit::refresh` is the same mutation inside a closure
  transaction (it composes with `set` and other edits).
- `refresh::Options`: `clock(Clock)`, `page_numbers(PageNumbers::{Keep,
  Require})`, `mark_dirty(bool)` (default on), `write_cache(bool)` (default
  off), and `calc(calc::Options)` with the `eval` feature.
- `refresh::Clock`: an explicit local date and time (`Clock::new`,
  `Clock::parse("YYYY-MM-DD[THH:MM[:SS]]")`, `FromStr`), and
  `Clock::format` for Word `\@` pictures. **The system clock is never
  read** (goal constraint 5): without a clock, date fields keep their cached
  result and are reported `needs_clock`.
- `refresh::Summary`: every field met, in document order, as
  `refresh::Field { kind, location, status, reason }`, plus `refreshed()`,
  `count(Status)`, `bookmarks()` (anchors added for TOC entries) and
  `missing_page_numbers()`. `Debug` prints counts only.
- `refresh::Status`: `Refreshed`, `Unchanged`, `Stale` (workbook cache
  differs and was not written), `NeedsLayout`, `NeedsClock`, `Locked`,
  `Unresolved` (missing bookmark or property), `Unsupported`.
- `capabilities::Verb::Refresh` (`"refresh"`, an edit): supported for DOCX
  and PPTX, for XLSX with `eval`, and report-only for XLSB/XLS with `eval`
  (`write_cache` refused there with `unsupported_capability`); every other
  format refuses with follow-up F1. Handles opened with
  `retain_source(false)` refuse as every edit does; signed and locked files
  refuse with `document_protected`.
- Atomicity and exactness: one refresh is one transaction operation and at
  most one native commit. A document whose computable results are current
  makes **no** native commit: the returned file is the input and
  `patch.is_noop()` holds (tested on Word- and LibreOffice-written TOCs, on
  refreshed output, and after save and reopen).

### DOCX

| Field | Result |
|---|---|
| `TOC` with `\o "a-b"`, `\u`, `\t "Style,n,..."` | rebuilt from the main story's headings: style outline levels (own or inherited through `basedOn`, an explicit body-text level stops inheritance), `HeadingN`/`heading N` styles, direct `w:outlineLvl` with `\u`, `\t` custom styles (by name, id, or a localized name equal to the id with spaces); `\h` links, `\n` page-number omission, `\p` separator; entries in styles `TOC1`..`TOC9` with a right dot-leader tab at the section's text width (US Letter's when undeclared) and a nested `PAGEREF` per entry |
| `SEQ` | document order per identifier: `\n`, `\c`, `\r N`, `\h`, `\* ARABIC/roman/ROMAN/alphabetic/ALPHABETIC` (OfficeCLI's evaluator); `\s`, other formats, or a failed earlier field stop that sequence (`unsupported`) |
| `REF` | the bookmark's text (`\h`, `\*` case formats); a missing bookmark is `unresolved` (Word's "Error! Reference source not found." is never written) |
| `DATE`, `TIME` | the clock, `\@` picture (default `M/d/yyyy`, `h:mm am/pm`) |
| `CREATEDATE`, `SAVEDATE`, `PRINTDATE` | the core properties' times |
| `TITLE`, `SUBJECT`, `AUTHOR`, `KEYWORDS`, `COMMENTS`, `LASTSAVEDBY`, `DOCPROPERTY` (those, `Category`, `Company`, `Manager`, `Template`) | the document properties; `\* Upper/Lower/Caps/FirstCap` |
| `NUMWORDS`, `NUMCHARS` | counted over the main story's paragraph text |
| `PAGE`, `NUMPAGES`, `SECTIONPAGES`, `SECTION`, `PAGEREF` | kept, `needs_layout` |
| anything else (`IF`, `HYPERLINK`, `MERGEFIELD`, `XE`, ...) | kept, `unsupported` |

Locked fields (`w:fldLock`) are `locked`. Fields nested in other fields
(except the `PAGEREF`s of a rebuilt TOC), in text boxes, or with `\#`
numeric pictures are `unsupported`, as is a field whose result holds
anything but plain runs.

**Page numbers (decision).** Litchi has no layout engine. OfficeCLI's
`refresh` regenerates TOC entries with page "0" and then needs Word (on
Windows) or a headless browser's HTML pagination; without either it fails
with `refresh_failed`. Litchi instead keeps each entry's cached page number
by matching the old entry's bookmark anchor (then its level and text), so a
document Word paginated keeps its numbers across heading edits. An entry
with no cached number gets an empty `PAGEREF` result, is counted in
`missing_page_numbers`, and the TOC's begin marker gets `w:dirty="true"`
(unless `mark_dirty(false)`) so Word recomputes it on open.
`PageNumbers::Require` refuses such a refresh with `refresh_failed`
(OfficeCLI's code), for callers who must not ship unknown numbers. `\n`
TOCs need no layout and are always complete.

**Bookmarks.** A heading's first `_Toc*` bookmark is reused; a heading
without one gets `_Toc<9 digits>` from a deterministic counter starting at
`_Toc100000001`, skipping names in use (OfficeCLI draws random GUID names;
Litchi output is byte-deterministic). Bookmark ids continue from the
largest numeric id.

**Exact no-op.** The old TOC's entries (level from the `TOCn` style id or
`toc n` name, text without the page suffix, anchor, page) are compared with
the planned ones; equal tables are `unchanged` and not rewritten. The three
LibreOffice-suite TOC fixtures (Word- and LibreOffice-produced) are exact
no-ops on both owners.

### PPTX

Slide-text `a:fld` fields: `slidenum` is `firstSlideNum` (read from the
presentation part; 1 when absent) plus the slide's position; `datetime` and
`datetime1`..`datetime13` use their en-US formats with the clock;
everything else is `unsupported`. One native text edit per changed field
run, all in one commit. Layouts and masters are not visited (their fields
are placeholders).

### XLSX (0266 follow-up F2)

`refresh` runs `File::recalculate` (record 0266). Without `write_cache` it
only reports: `unchanged` (match), `stale` (mismatch or uncached),
`needs_clock` (volatile), `unsupported`. With `write_cache`, every stale
result is written as the cell's cached value — the `<f>` element is kept
byte for byte (shared-formula members included), the `<v>` and the cell
type (`str`, `b`, `e`, number) follow the value — in **one** native
workbook edit (exactly once: a second refresh is an exact no-op). Values a
workbook cannot hold (non-finite numbers, error values without an XLSX
spelling) and non-scalar formulas are `unsupported` and left untouched.

### Leaf primitives (small, separate hunks)

- `litchi-docx` (`document/transaction/fields.rs`, new, plus forwarding in
  `document/owned/fields.rs`): `document::fields::{Fields, Paragraph,
  Field, Bookmark, Update, TocEntry, Outcome}` and `Edit::{fields,
  refresh_fields}`. The scan is namespace-aware and bounded (262,144
  fields, nesting 64, bookmark text 64 KiB, the transaction's depth and
  byte limits); `mc:Fallback` content is skipped. All updates of one call
  are spliced once; changed body children are recorded as exact
  `ReplaceElement` operations, and a TOC whose paragraphs are direct body
  children as exact `RemoveElement`/`InsertElement` operations (applied last
  to first), so the patch replays and inverts exactly (unit-tested). A text
  update keeps the instruction, markers and the first result run's
  formatting; a TOC update keeps the begin paragraph through its separator
  and the end paragraph from its end marker (a simple-field TOC becomes a
  complex field). `element.rs` gains `body_children` (crate-private);
  `styles::Style::declares_outline` exposes an explicit body-text
  `w:outlineLvl`, which stops outline inheritance.
- `litchi-xlsx`: `WorksheetEdit::set_cached` and the private
  `Payload::Cache` (validation, projection, patch verification and the
  sheet-data writer), refusing non-formula and non-scalar cells.
- `litchi-pptx`: run-text edits accept an `a:fld` target (inserting `a:t`
  when the field has none); `Package::first_slide_number` and
  `source::Package::first_slide_number`.

### Surfaces

- CLI: `litchi refresh FILE [--clock DATETIME] [--write-cache]
  [--require-page-numbers] [--no-dirty] [--max-changes N] [--max-bytes N]
  [--out F]`. It saves in place (OfficeCLI's `refresh` modifies the file),
  or to `--out`, and only when a result changed. Text output starts with
  `Refreshed: <path>` (OfficeCLI's line) or `Up to date: <path>`, then one
  line per field; `--json` prints `output`, `changed`, `refreshed`,
  `bookmarks`, `missingPageNumbers` and `fields` (`type`, `path`, `status`,
  `reason`).
- Python: `File.refresh(*, clock=None, write_cache=False,
  page_numbers="keep", mark_dirty=True, max_changes=None, max_bytes=None)
  -> Commit`, GIL released; `clock` is a str, `datetime.datetime` (wall
  time to the second) or `datetime.date`; `Commit.value` is
  `_RefreshSummary` (stubs updated, `test_refresh.py`).

## 2. Deltas vs OfficeCLI

OfficeCLI reference: `CommandBuilder.Refresh.cs`, `Core/WordTocBuilder.cs`,
`Core/WordHtmlRefresh.cs`, `Handlers/Word/WordHandler.SeqEval.cs`
(read-only).

| Id | OfficeCLI | Litchi | Disposition |
|---|---|---|---|
| R1 | `refresh` is `.docx` only; other formats `unsupported_type` | DOCX, PPTX slide fields, XLSX caches (and XLSB/XLS reports) | superset |
| R2 | page numbers from Word (Windows) or a headless browser; none available is `refresh_failed` | cached numbers kept by anchor; unknown ones empty plus `w:dirty`, or `refresh_failed` under `--require-page-numbers` | deliberate (no layout engine, no ambient processes) |
| R3 | headings: direct `outlineLvl` first, then the style's own level, then `HeadingN` id, whatever the switches | Word's switch semantics: style levels for `\o`, direct levels only with `\u`, `\t` custom styles, `basedOn` inheritance, explicit body text stops it | deliberate |
| R4 | `\t`, `\n "a-b"`, `\p` not interpreted (`\n`/`\z` both suppress page numbers) | interpreted; `\z` (web view) only kept in the code | superset / fix |
| R5 | new `_Toc` bookmarks with random names | deterministic `_Toc<n>` names | deliberate (determinism) |
| R6 | the TOC's end marker moves to a fresh paragraph; text after it in the old end paragraph is dropped | the end paragraph keeps its properties and everything from the end marker on | fix |
| R7 | entries without tab stops or `TOCn` style definitions | entries carry a right dot-leader tab at the text width; `TOCn` style definitions are not added to the styles part (Word supplies built-ins) | partial; F4 |
| R8 | SEQ evaluated at write time (`RecalcSeqFields`), not by `refresh` | by `refresh`, same switch scope | parity in result |
| R9 | `refresh` rewrites unconditionally | exact no-op when current; nothing saved | deliberate |
| R10 | REF/DATE/TITLE/... left to Word | computed (DATE/TIME only with `--clock`) | superset |
| R11 | `ExtendedFilePropertiesPart.Pages` updated from pagination | not touched | wontfix without layout |

## 3. Tests

- DOCX leaf (`fields/tests.rs`, 8): index contents (styles, outline
  levels, bookmarks with text, text width, field results and lock state),
  text updates in the three result shapes with formatting kept, no-op text
  updates, a body-level TOC rebuilt with bookmarks and dirty marking, a
  multi-paragraph TOC in a content control, a simple-field TOC made complex,
  refusals without change, stale index and bad text refusals. Every commit
  is replayed onto its source and inverted exactly.
- XLSX leaf (`formula_cache_writes_keep_formula_markup`): shared-formula
  members and string results keep their `<f>`, types follow values, a
  non-formula cell is refused, the inverse restores the bytes; doctest.
- PPTX leaf (`field_results_take_run_text`).
- Ops: `refresh/tests.rs` (instruction parsing, clocks and pictures, number
  formats and cases), `format/docx/refresh/tests.rs` (SEQ semantics, names
  and ranges), integration `tests/refresh.rs` (8: fixture no-ops on both
  owners, a changed heading keeps page numbers and reopens current, the
  OfficeCLI placeholder TOC and every field kind with statuses, the clock,
  `Require` and `mark_dirty(false)`, a transaction composing two refreshes
  with an edit, PPTX slide number and `datetime1`, refusals and the
  capability), `tests/xlsx_refresh.rs` (2: report, write-back of number,
  string, boolean and error results in one commit, exactly-once after
  reopen), the `File::refresh` doctest, and the capability goldens.
- CLI `refresh_saves_only_changed_documents`; Python `test_refresh.py` (3).
- LibreOffice 26.2.6.3 headless (`soffice --convert-to txt`) opens the
  refreshed Word fixture and a refreshed OfficeCLI-style placeholder
  document and renders the new entries (`Chapter Two	2`; `Introduction`,
  `Background`, `Appendix`), the SEQ, REF and DATE results.

Fixtures (copied into `test-data/libreoffice-core/`, recorded in its
`PROVENANCE.md` with SHA-256, MPL-2.0): `sw/qa/extras/ooxmlexport/data/
{multi-page-toc,testTOCFlag_u,custom-styles-TOC-comma}.docx`,
`oox/qa/unit/data/testTdf132557_footerCustomShapes.pptx`.

## 4. Corpus and capability delta

`LITCHI_BLESS=1 cargo test -p litchi-cli --all-features --test
officecli_corpus`, `python3 tools/officecli_corpus.py`, `--check`: no
diff. No corpus script calls `refresh` (`word/fields.sh` relies on
`updateFields=true` instead), so outcomes are unchanged. Capability goldens
were re-blessed with all features and with every feature except
`encryption`; the only change is the new `refresh` verb. The OfficeCLI
capability delta table compares the property registry, which refresh does
not touch.

## 5. Performance

`cargo build --release -p litchi-ops --features docx,xlsx,eval --example
refresh_bench; taskset -c 29 target/release/examples/refresh_bench 1000 11`
(2 warm-up legs, 11 ABBA rounds of four samples, medians; machine load
average about 17 on 32 cores, six agents). Evidence:
[results/0279-refresh-bench.json](results/0279-refresh-bench.json) (3.9 KB,
every sample, binary SHA-256 recorded inside). Both legs' outputs are
checked equal outside the clock (document XML; workbook bytes).

| Scenario | Leaf ms | Unified ms | Unified / leaf | Unified items/s |
|---|---:|---:|---:|---:|
| TOC over 1,000 headings, built (1,000 bookmarks added) | 4.75 | 4.86 | 1.022 | 205,885 headings |
| the same, already current (exact no-op) | 6.97 | 6.16 | 0.885 | 162,237 headings |
| XLSX, 6,000 uncached formulas written back | 34.67 | 34.85 | 1.005 | 172,191 formulas |

- The leaf TOC leg is the hand-written equivalent over the new leaf API
  (edit, field index, entries planned knowing the fixture's `HeadingN`
  styles, `refresh_fields`, commit); the unified leg adds the style part,
  the instruction grammar, page-number matching, the summary and the
  transaction patch. In the current case the unified planner proves the
  table current and makes no edit, while the leaf leg's rewrite only finds
  its result unchanged at the end, hence < 1.
- The first measurement had the current case at 1.33: the planner parsed
  every nested `PAGEREF` instruction and matched page numbers with a linear
  search per entry (quadratic). Lazy instruction parsing, hashed page
  lookups and building entries only for a changed table brought it down.
- XLSX: both legs share `File::recalculate` (about 90% of the time); the
  write-back itself is one native edit of 6,000 `set_cached` calls.
- All three are within the goal's 10% single-operation budget.

## 6. Gates

All with `CARGO_TARGET_DIR=<worktree>/target`, exit 0:

- `cargo fmt --all --check`;
- `cargo clippy -p litchi-docx -p litchi-xlsx -p litchi-pptx -p litchi-ops
  -p litchi -p litchi-py -p litchi-cli --all-targets --all-features -- -D
  warnings`, and `cargo clippy -p litchi-ops --no-default-features
  [--features F] --all-targets -- -D warnings` for no features, `docx`,
  `pptx`, `xlsx`, `xlsx,eval`, `xlsb,eval`, `xls,eval`, `eval`, `rtf`;
- `cargo test --all-features` for `litchi-docx` (2,380 passed),
  `litchi-xlsx` (2,445), `litchi-pptx` (1,773), `litchi-ops` (2,123,
  `--no-fail-fast`), `litchi` (260), `litchi-cli` (31);
- `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features` for the six
  Rust crates above;
- `python3 tools/check_crate_boundaries.py`, `python3
  tools/non_iwork_gate.py verify`, `RUSTC_BOOTSTRAP=1 python3
  tools/check_api_glossary.py --target-dir target/api-glossary` (26
  profiles);
- `crates/pyo3-litchi/tests/run_tests.sh` (264 tests);
- `LITCHI_BLESS=1 cargo test -p litchi-cli --all-features --test
  officecli_corpus`, `python3 tools/officecli_corpus.py` and `--check` (no
  diff), `python3 tools/officecli_capability_delta.py --officecli …` (no
  diff);
- `cd tools/perf-baseline && cargo check --locked --all-targets`.

Two existing tests walk every fixture under `test-data/`, so the four new
fixtures moved them:

- `litchi-docx` `the_route_and_proof_follow_the_historical_gate_on_mutated_sources`
  counts fallbacks over the first 20 fixtures in path order; the new files
  sort into that window and left the count at its bound (20, needs > 20).
  It now takes the first 40.
- `litchi-ops` `query_corpus_differential` (record 0262): with the four new
  files moved aside the digests match the golden unchanged; with them, the
  golden was re-blessed (`LITCHI_BLESS_GOLDEN=1`), changing only the
  aggregate DOCX digests the new files enter.

## 7. Gaps and follow-ups

- F1 refresh for DOC, PPT, RTF and ODF (field results exist in their
  leaves; no refresh adapter yet).
- F2 DOCX headers, footers, notes and text boxes are not refreshed (main
  story only); PAGE/NUMPAGES there need layout anyway.
- F3 page numbers: no layout engine; an HTML-pagination or layout source
  could fill them (OfficeCLI's fallback). `docProps/app.xml` `Pages` is not
  updated.
- F4 `TOC1`..`TOC9` and `TOCHeading` style definitions are not added to a
  styles part that lacks them (Word and LibreOffice supply built-ins).
- F5 TOC switches `\a`, `\b`, `\c`, `\d`, `\f`, `\l`, `\s` (captions, TC
  entries, bookmark scope, chapter numbers) are `unsupported`; `\w`, `\x`,
  `\z` are kept but have no effect without layout.
- F6 `\#` numeric pictures; `SEQ \s`; `REF \n/\r/\w/\p` (paragraph numbers
  and relative positions); `DOCPROPERTY` custom properties; `NUMPAGES`-like
  statistics from `app.xml`.
- F7 PPTX: notes, layouts and masters; `datetime` formats are en-US only.
- F8 XLSX: shared strings are not used for string results (`t="str"` inline
  value, as Excel writes); XLSB/XLS cache write-back needs leaf cached-value
  writes.
- Finding for record 0273 (find/replace): a DOCX find over `/body` searches
  the result of a table of contents whose field spans paragraphs (its
  per-paragraph field tracking resets at paragraph boundaries), so a find
  can rewrite TOC entries; the refresh tests address the heading paragraph
  instead.
