# 0338 — XLSX P0 residue, root-level metadata writes, exact JSON numbers and a PathArg display fixed point (WP16; record 0334 F1, F2)

- Status: implemented on `uo2/xlsx-p0-meta`, cut from `feat/unified-ops-api`
  at `d165759e4a` (record 0334 merge). Trunk was merged at `f72374e314`
  (records 0333 and 0335).
- Date: 2026-10-17
- Input: record 0332 §7 (WP16, proposed), record 0334 findings F1 and F2,
  and the definition-of-done audit of 2026-10-17, §2. In the v1 corpus
  subset, 38 lines were blocked on WP16.
- Owned:
  - `crates/litchi-xlsx` (rich-text run fonts, the source-backed rich read,
    picture image replacement);
  - `crates/litchi-ops/src/format/xlsx*`;
  - the root-metadata routing in the DOCX, XLSX and PPTX adapters, in small
    hunks (`format/metadata.rs`, `format/{docx,pptx,xlsx}/metadata.rs`, one
    hunk each in `format/docx/write.rs`, `format/docx/write/coalesce.rs`,
    `format/docx.rs` and `format/pptx.rs`);
  - `crates/litchi-ops/src/path` (F2, and the `r` run alias);
  - the workspace `serde_json` features (F1);
  - `crates/litchi-ops/fuzz`.
- Small hunks outside them:
  - shared dispatch: `capabilities.rs` (the settable, addable and
    removable claims; the XLSX workbook capability table) and
    `transaction.rs` (a DOCX metadata step's change test);
  - `schema.rs` (the run aliases);
  - `litchi-drawingml` (`Spec::with_series_reference`, additive): the
    chart authoring model had no per-series reference.
  - `litchi-ops/Cargo.toml` (`proptest` dev-dependency).
- `performance_claim`: measured (§6). Three new paths miss the 10%
  single-op budget, explained in §6.

## 1. Root-level metadata writes (DOCX, XLSX, PPTX)

OfficeCLI writes document properties on the root: `set / title=…`.
Litchi wrote them on `/metadata` only (records 0253, 0254), and the root
refused (`unsupported_element` for DOCX and XLSX, a warned skip of every
key for PPTX). The root now routes to the same writer, with the same keys
(the 22 OfficeCLI root-metadata spellings and `custom.<name>`):

| Format | Routed | Other root keys |
|---|---|---|
| XLSX | every non-empty root `set` (the root sets nothing else) | `unsupported_property`, named |
| PPTX | a root `set` naming at least one metadata key | `unsupported_property`, named; a root `set` without a metadata key keeps its old route |
| DOCX | every non-empty root `set`, and now `/metadata` too | `unsupported_property`, named |

- **DOCX** had no unified metadata writer. It now commits the leaf's
  core-properties transaction (`litchi_docx::metadata`, record 0251), so
  the core keys are written: title, author, subject, keywords,
  description, category, lastModifiedBy, revisionNumber, created,
  modified, contentStatus, language, version, lastPrinted. Application keys
  (`extended.*`, `company`, `manager`, `template`, `application`),
  `contentType` and `custom.<name>` are refused by name: the leaf family
  wraps the core part only (record 0253 F3).
- Under `Unsupported::Warn` the other root keys are skipped and the
  metadata is written (`tests/root_metadata.rs`).
- The root elements now claim `set`, and their document-property keys are
  settable (`root_capability` over the read table, same order). So the
  record 0330 consistency test holds, and `capabilities()` agrees with the
  registry. XLSB keeps `/metadata` only.
- Transactions: a DOCX metadata step records an empty main-document patch
  (the handle's own editor committed empty). The coordinator's change test
  also compares the endpoints' core properties, so the step is not a no-op
  (`transactions_replay_and_invert_root_sets`).
- `File::set` on a DOCX root or `/metadata` does not build a main-document
  editor. It applies the same policy (`Docx::edit_policy`: signed,
  `retain_source(false)`, authorized signature removal) before the leaf
  transaction (§6).

## 2. XLSX P0 residue

### 2.1 Rich-text runs (`run`)

The run element is `/<sheet>/<A1>/run[N]`, with OfficeCLI's `r[N]` as an
alias. It has OfficeCLI's keys (`RunToNode`): `bold`, `italic`, `strike`,
`underline`, `superscript`, `subscript`, `size`, `color` and `font`. Their
aliases are `font.bold`, `font.italic`, `font.strike`, `font.underline`,
`font.size`, `fontsize`, `font.color`, `font.name` and `fontname`. `text` is
the node text, and a write-only key.

- **Leaf.** `RichRun::font` reads a run's `rPr` as the stylesheet `Font`.
  `RichRun::apply_font` applies a `FontChange` in place: a property the
  change names is replaced or removed, every other property keeps its
  markup, and a new property goes in Excel's order. Also new:
  `RichRun::{new, set_text}`, `RichText::{new, run_mut, insert_run,
  remove_run}`, and `SourceWorksheet::rich_text` (the source-backed read).
- **Reads.** `get` of a run path, and `query run`, read the cell's runs:
  an inline rich string, or a shared-string item with runs. A walk visits
  a cell's runs only when it asks for runs, and only for text cells. A
  cell's own view is unchanged: it lists no run children (delta R1).
- **Writes.**
  - `set` changes one run.
  - `add` under a cell inserts a run, at the end or at a placement. A plain
    text cell becomes rich, and its text becomes the first run, as in
    OfficeCLI's `AddRun`.
  - `remove` deletes a run. As in OfficeCLI, if no run is left the cell
    holds empty text, and a single unformatted run left becomes plain
    text.
  - Every write publishes the runs as the cell's inline rich string, as
    find/replace does (record 0295).
- **Cell `runs`.** On `add` and `set`, `runs` is OfficeCLI's JSON array
  of run objects (`ApplyRichTextToCell`). It works with or without
  `type=richtext`. Each object's keys go through the run registry. An
  unknown key, an empty array, or `runs` given together with a value is
  refused.

### 2.2 The cell `type`

`type` takes `string`, `number`, `boolean`, `date`, `error` and
`richtext`, plus the aliases `str`, `num`, `bool`, `err` and `rich`. It
works on `set` and `add`:

- with a value, it converts the value: a text that is no number is
  refused for `number`; `boolean` accepts true/false, yes/no and 1/0; an
  error value is one of the eight Excel tokens;
- alone, it converts the value already stored in each target. An empty
  coordinate is left alone, and a formula is refused;
- `richtext` with a text value writes one unformatted run;
- combined with a formula or `clear`, it is refused.

`type=date` stores an ISO 8601 `t="d"` cell (`Value::Date`), not a serial
number (delta R2).

### 2.3 `add` of a cell

The parent picks the cell, as OfficeCLI's `AddCell` does:

- under a sheet, `ref` (alias `address`) names the cell; without one, the
  first empty cell of row 1 is used;
- under a row, `ref` names the cell; without one, the row's first empty
  cell is used;
- under a cell path, that cell is used.

The other properties are the cell `set` properties, written in one native
edit, and the result path is `/sheet[N]/<A1>`.

- An `add` needs at least one property.
- It never overwrites a cell that holds a value or formula (delta R3).
- A `ref` that disagrees with a cell or row parent is refused (delta R4).
- The cell element now claims `add`, and its payload and formatting keys
  are addable. `ref` is add-only. A range claims no `add`, so its
  capability rows keep `add` false.

### 2.4 Picture source replacement

`set src` (alias `path`) on `/<sheet>/picture[N]` replaces the image, with
the same sources `add` reads (a file, or a data URI of a PNG, JPEG, GIF,
BMP or TIFF image). The new leaf verb is
`drawing::edit::Edit::replace_image`:

- it writes the bytes as a new media part, retargets the picture's own
  image relationship, and removes the old part once nothing refers to it;
- the anchor, names and geometry keep their bytes;
- the same image is an exact no-op;
- a picture whose image relationship another object shares is refused,
  because replacing it would change both.

A `src`-only set does not rewrite the anchor. `src` is now `set` and `add`
in the registry.

### 2.5 Chart series from another sheet

On a chart `add`, `series{N}.values` may now be a cell range on any sheet:
`Sheet1!B2:B13`, `'My Data'!$C$1:$C$4`, or `B2:B5` on the chart's sheet.

- The values are read from the cells, and a cell without a number counts
  as 0, as with `dataRange`. The range is one row or one column.
- The series refers to the range through an absolute reference, with
  cached values (`Spec::with_series_reference`).
- This works for category charts: column, bar, line, area, pie, doughnut,
  radar and combo. A scatter or bubble chart, or a chart with
  `dataRange`, still refuses a range (`unsupported_capability`, as
  before).
- `series{N}.categories` is still not a key (delta R5). The corpus
  charts now warn on it instead of failing.

## 3. F1: exact JSON numbers

The workspace's `serde_json` is now built with `float_roundtrip`:

```toml
serde_json = { version = "1", features = ["float_roundtrip"] }
```

Every crate's JSON number parsing is now correctly rounded.

- `98446744073709551615` decodes to `9.844674407370955e19`, Python's value
  (record 0334 decoded `…957e19`).
- The canonical batch wire form is a fixed point: it re-serializes to the
  same bytes. This covers the 121-digit literal of record 0334
  (`batch::tests::numbers_decode_correctly_rounded_and_round_trip_exactly`).
- The fuzz harness drops its 4-ULP tolerance (`same_items`). It now
  requires equal items and byte-identical re-serialization.

**Intended output change.** A `PropValue::Number` decoded from JSON can
differ from the old value by one or more ULPs. That changes the decoded
value, not the bytes Litchi writes for it. Only long, or not exactly
representable, decimals are affected. No golden, corpus outcome or
fidelity row changed.

**Cost** (`examples/json_number_bench.rs`; two release builds, one per
side; `taskset -c 7`; 6 ABBA blocks; each run is the median of 31 decodes;
[results/0338/json-number-bench.txt](results/0338/json-number-bench.txt)):

| Payload | before | after | after/before |
|---|---:|---:|---:|
| OfficeCLI seeds (30 items, 5 KB, numbers as strings) | 10.1 µs | 10.7 µs | 1.063 |
| 10k cell sets, short decimals (`12.5`) | 3.70 ms | 3.82 ms | 1.033 |
| 10k cell sets, 17-digit decimals with exponents | 3.87 ms | 4.49 ms | 1.159 |
| 10k cell sets, integers | 3.54 ms | 3.69 ms | 1.044 (no float parse; noise and code layout) |

Long decimals cost about 30 ns more per number, which is the price of a
correctly rounded parse. The other rows are within the host's noise
(ranges in the evidence file). The checksums of the decoded numbers
differ only in the long-decimal row, where the old parse was not
correctly rounded.

## 4. F2: `PathArg` display is a fixed point of `parse`

The core grammar writes a relative path `'bod! my'` bare, as `bod! my`.
`PathArg::parse` read that spelling back as Excel sheet notation, which
is another node (`/bod/ my`).

`PathArg`'s `Display` now quotes the first step of a relative path when
its bare spelling would read as sheet notation:

- the name has no leading quote and no `/`, `[`, `]` or control
  character;
- it holds a `!` that is not its first byte;
- the `!` has no `"` before it and no `=` after it.

So `'bod! my'` now displays as `"bod! my"`.

- Every other spelling is unchanged, and so is parsing.
- Absolute paths and sheet notation keep the core spelling.
- A proptest (`path::arg::display_tests`) checks the fixed point over text
  built from the bytes that matter (`a 1 ! ' " / = [ ] : @ \`), and over
  typed relative and absolute paths. A unit test pins the cases.
- The `paths` and `batch_payload` harnesses now reparse through
  `PathArg::parse`, and require the same path and spelling. The
  core-grammar workaround is gone.

## 5. Corpus, delta and checklist

**v1-scope subset** (`officecli-corpus-deltas.md`):

| Tree | Lines | Passed | Accounted | Open | Open on WP16 |
|---|---:|---:|---:|---:|---:|
| base `d165759e4a` | 3,443 | 1,877 | 941 | 625 | 38 |
| trunk `e39d0a75a4` (records 0333, 0335) | 3,443 | 2,184 | 1,062 | 197 | 38 |
| this branch after the merge | 3,443 | 2,204 | 1,069 | 170 | 0 |

What happened to the 38 WP16 lines:

- **20 pass:**
  - the nine `type=string` sets;
  - the six cell adds (`cell-formatting.sh` and `charts-bubble.sh`);
  - the five root metadata sets (two XLSX, two PPTX, one DOCX).
- **7 are accounted (deferred-P1):**
  - the three cross-sheet chart adds now warn only on
    `series{N}.categories`;
  - four DOCX root sets warn only on P1 document settings
    (`mirrorMargins`, `embedFonts`, `compatibility.mode`,
    `updateFields`).
- **10 are open as unverified.** These are root sets of XLSX, PPTX and
  DOCX document settings: the calc, protection, theme, page-setup,
  docDefaults and docGrid keys. Their warned lists are longer than the
  corpus outcome's 48-character detail bound (record 0332 F1, the corpus
  harness in `litchi-cli`). The listed keys are deferred-P1.
- **1 is open, planned 0253.** `word/document-formatting.sh:28` sets
  `extended.*` on a DOCX root, which needs DOCX application properties
  (record 0253 F3).

**Dispositions.** The WP16 entries are closed and deleted:
`metadata-root`, `docx-root-writes`, `xlsx-p0-residue` and
`xlsx-cell-add`. Both `--check` runs are green.

**Capability delta** (trunk at the merge → this branch):

- XLSX elements: supported 33 → 34 (the run).
- XLSX property reads: supported 333 → 340.
- Property writes, supported:
  - XLSX 212 → 268 (partial 60 → 26);
  - DOCX 260 → 266;
  - PPTX 142 → 153.
- Planned property gap rows:
  - XLSX 23 → 0;
  - DOCX 85 → 79;
  - PPTX 27 → 16.

**CRUD checklist** (`results/0289/crud-checklist.*`):

- New cells run and pass: XLSX `cell add` (new recipe: `ref` and `value`
  on `create()`); XLSX `run` get, find, query, view, set, add, remove,
  edit and batch; root `set`, `edit` and `batch` for DOCX `document`,
  XLSX `workbook` and PPTX `presentation`.
- Totals: DOCX 305 ok, XLSX 206, PPTX 282, with 0 FAIL.

**Fidelity** (`results/0289/fidelity.*`): open, then empty edit, then save
keeps every byte (300 of 300 rows exact, unchanged). The harness probes
the first settable property of each document, which is now a root
document property (`/ author`, `/ title`). So 122 single-set rows moved
from byte-exact to value-exact, because reverting a metadata set rewrites
`docProps`. Another 14 rows refused the old probe and now pass the new
one: 7 DOCX (protected styles), 5 XLSX and 2 PPTX.

**Goldens.** Re-blessed in both feature sets. The changed rows are:

- the root elements' `set` claims and their document-property keys;
- the XLSX run element (props, no gap);
- the cell `add` claim and the cell key `add` flags;
- the new cell keys `type` (set/add), `runs` and `ref`;
- picture `src` set.

## 6. Measurements

`examples/p0_residue_bench.rs`, release, `taskset -c 7`, eager handles,
ABBA, the median of 41 rounds, with outputs compared outside timing. The
leaf leg is the hand-written edit doing the same work, from the same
snapshot. Final run, after the trunk merge:
[results/0338/p0-residue-bench.txt](results/0338/p0-residue-bench.txt).

Run 1 of the final file (run 2 agrees within 1%):

| Scenario | Leaf | Unified | Ratio |
|---|---:|---:|---:|
| DOCX `set / title` | 12.8 µs | 13.2 µs | 1.03 |
| XLSX `set / title` | 8.9 µs | 9.2 µs | 1.04 |
| XLSX `set /metadata` vs `set /` (same writer) | 9.0 µs | 9.0 µs | 1.00 |
| PPTX `set / title` | 73.1 µs | 73.6 µs | 1.01 |
| XLSX cell `value` + `type=number` | 27.3 µs | 28.6 µs | 1.05 |
| XLSX cell `value` (existing hot path) | 27.5 µs | 28.7 µs | 1.05 |
| XLSX `add` cell (`ref`, `value`) | 27.6 µs | 29.5 µs | 1.07 |
| XLSX run `set bold` | 1.85 µs | 2.45 µs | **1.32** |
| XLSX picture `set src` (both decode the data URI) | 16.4 µs | 20.7 µs | **1.27** |

- The first version of the DOCX root set was 1.40. It built a
  main-document editor, whose markup-compatibility scan of the main part
  callgrind showed at 13%, and committed it empty for the patch. The
  single-op route now skips it. A transaction step still builds the
  editor, because it needs the patch.
- **Misses.**
  - The run set is 0.55 µs over a 1.9 µs leaf edit. That is the fixed
    per-call cost: path parsing, sheet resolution, the cell view, and
    property admission. The leaf leg reads the runs once, and so does the
    unified one, since it reuses the runs its path resolution read.
  - The picture set is 4.3 µs over the leaf. Path resolution reads the
    worksheet's drawing elements to find `picture[1]`, and
    `replace_image` reads them again. Follow-up 0338-P1: a leaf
    `replace_image` that takes the resolved element.
- **Cell writes.** The cell registry grew from 51 to 54 entries, and the
  set route gained the run-path check. The plain cell set stays at 1.05,
  the same ratio as before this change (records 0097, 0330).

## 7. Deltas against OfficeCLI

These are deliberate differences, recorded here. None is visible in the
corpus as a failure.

- **R1.** A cell's view does not list its runs as children, and its
  `childCount` stays 0. OfficeCLI decomposes SST runs into `/run[N]`
  children of the cell's `Get`. In Litchi, runs are addressed by path and
  by `query run`, so bulk cell listings read no shared-string runs. A
  follow-up would list them for a cell resolved by path only.
- **R2.** `type=date` stores an ISO `t="d"` cell. OfficeCLI stores an
  OADate serial number with a date number format.
- **R3.** `add` of a cell that holds a value or formula is refused. It
  is `set`'s job, as `add` of an existing style is refused (record 0319
  D11). OfficeCLI overwrites.
- **R4.** A `ref` that disagrees with the parent path is refused.
  OfficeCLI warns, and uses `ref`.
- **R5.** `series{N}.categories` is not a chart key: the shared chart
  registry lists `name`, `values` and `color`. It is deferred-P1 by the
  record 0332 breadth rule.
- **R6.** Runs are written as the cell's inline rich string. OfficeCLI
  appends a shared-string item. Both are valid, and readers see the same
  text and runs.
- **R7.** A PPTX root `set` that mixes a metadata key with another root
  key refuses the other key, so `Warn` skips it. Without a metadata key,
  the root keeps its previous route.

## 8. Not claimed, and follow-ups

- **DOCX application properties** on the root and `/metadata`
  (`extended.*`): record 0253 F3, still planned.
- **Root document settings** (calc, theme, page setup, docDefaults and so
  on) are not written. They are `unsupported_property`, deferred-P1.
- **The corpus detail bound** (record 0332 F1) hides 96 open lines, 10 of
  them from this record.
- **0338-P1.** Picture `set src` reads the drawing elements twice (§6).
- **0338-P2.** `seriesN.values` ranges for scatter and bubble charts, and
  `set` of a series' values to a range: the shared chart edit still
  refuses ranges.
- **Run move and swap** are not offered (the checklist shows
  "format-verb, unreached").
- The F1 cost was measured on the batch decoder only. The other JSON
  readers (CLI batch files, Python `apply`) use the same parser.

## 9. Verification

The gate commands and their exit codes are in this branch's final report:
`fmt`; `clippy` (all features, every litchi-ops format feature, and
without encryption); the `litchi-xlsx`, `litchi-drawingml`, `litchi-ops`
(in chunks), `litchi`, `litchi-core` and `litchi-cli` tests; `rustdoc`;
`pyo3`; the feature matrix; fuzz `check` and the smoke runs; `rebless` and
`golden-clean`; `perf-baseline` and `native-resave` `check --locked`; the
glossary, crate-boundary and non-iWork gates; the tool unit tests; the
bench matrix; and both dispositions `--check` runs.
