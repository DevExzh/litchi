# 0320 — XLSX P0/P1 worksheet features (OfficeCLI parity, tier 1)

- Status: implemented on `uo2/xlsx-p0p1`, cut from `feat/unified-ops-api`
  at `f94b388742`.
- Date: 2026-10-16
- Input: `officecli-capability-delta.md` XLSX element gaps L-XLSX-R5
  (comments), L-XLSX-R7 (pictures, shapes) and M1-extended (R4x)
  (hyperlinks, conditional formatting and its 12 rule kinds, validation,
  autofilter, row/column/page breaks); design §13 P0.
- Owned: `litchi-xlsx` (new edit families, reads, one codec fix), the XLSX
  adapter of `litchi-ops` (`format/xlsx/{features,feature_write}.rs` and
  small hunks in `node`, `nav`, `children`, `spec`, `write`, `structure`,
  `sheet_write`), the XLSX rows of the registry, capability, checklist and
  corpus goldens, one Python test.
- `performance_claim`: none beyond §6, which reports the paired
  measurements of the new paths.

## 1. Result

| OfficeCLI element | before | after |
|---|---|---|
| `comment` (notes) | gap L-XLSX-R5 | get, query, add (sheet + `ref`, or a cell parent), set (`text`, `author`, `ref` moves), remove |
| `hyperlink` | gap | get, query, add, set, remove; cell `link`/`tooltip`/`display` on `set` (`link=none` removes) |
| `picture` | gap L-XLSX-R7 | get, query, add (`src` data URI or path; `x`/`y`/`width`/`height` in cells or lengths, `anchor=B2:E6`, `anchorMode`), set (geometry, `name`, `alt`, `rotation`, flips), remove |
| `shape` | gap L-XLSX-R7 | get, query, add (preset `geometry`, `fill` RGB/scheme/`none`, `line` color:width:dash, text and run formatting, `align`, `valign`, `margin`, `rotation`, flips, `anchor`), set (geometry, `name`, `alt`, `rotation`, flips), remove |
| `conditionalformatting` + `cellis`, `formulacf`, `colorscale`, `databar`, `iconset`, `topn`, `aboveaverage`, `duplicatevalues`, `uniquevalues`, `containstext`, `dateoccurring`, `cfextended` | gap | get, query, add (`--type cf` with `type=`, or the kind word directly), set, remove; formatting through a new differential format |
| `validation` | gap | get, query, add, set, remove (overlap refused, as OfficeCLI) |
| `autofilter` | gap | get, query, add, set (`range`), remove; sheet `autoFilter=A1:B5`/`none` |
| `rowbreak`, `colbreak`, `pagebreak` | gap | get, query, add, set, remove (`pagebreak` dispatches on `col`) |
| `run`, `sparkline` (optional) | gap | gap (not done, §7) |

Delta table (`officecli-capability-delta.md`, XLSX): elements supported
**11 → 33**, gap **27 → 5** (ole, pivottable, run, slicer, sparkline);
property writes supported **54 → 211**.

OfficeCLI corpus (`officecli-corpus-deltas.md`), family `excel`: passed
**1484 → 1552**, warned **21 → 26**, gap **124 → 51**, bug 0; family
`excel/charts` unchanged (warned 255, gap 22). The five new warned calls are
calls that were refused whole before and now apply, skipping only the
effect or x14 properties §7 lists.

## 2. Leaf: `litchi-xlsx`

One staged engine (`part_edit`, crate-private) backs five scoped families
in the ADR 0034 shape, each registered in the glossary gate
(`tools/api_glossary.json`, profiles `xlsx-minimal` and `xlsx-full`):

| family | entry points | edits |
|---|---|---|
| `conditional_formatting::edit` | `Package::{edit_conditional_formats, apply_conditional_formats}` | `formats`, `set_formats`, `add_differential(&DifferentialStyle)` |
| `data_validation::edit` | `Package::{edit_data_validations, apply_data_validations}` | `validations`, `set_validations` |
| `auto_filter::edit` | `Package::{edit_auto_filter, apply_auto_filter}` | `auto_filter`, `set_auto_filter` |
| `workbook::comments::edit` | `Package::{edit_comments, apply_comments}` | `comments`, `set`, `add`, `remove` |
| `drawing::edit` | `Package::{drawing_elements, edit_drawings, apply_drawings}` | `elements`, `add_picture`, `add_shape`, `set(Change)`, `remove` |

`Staged` clones the package's OPC graph, records every part it touches, and
`commit` captures the exact before/after state of those parts (content
type, payload, relationships) as the reversible patch; an unchanged staging
returns the source package, so an exact no-op keeps every byte. Patches
replay with `apply_*` onto an equal package and refuse any other with
`PatchConflict`; a changed edit of a signed package is `Signed`.

- **Notes** are splices: one `<comment>` (and author) is inserted, replaced
  or drained; every other note keeps its bytes, rich text and `xr:uid`
  included (the existing `write_comments` flattens rich text, so it is not
  used on existing parts). A new note gets the hidden VML note shape Excel
  draws it from; the comments part, the VML part and `<legacyDrawing>` are
  created with the first note and removed with the last one when the VML
  holds nothing else.
- **Drawings**: pictures stage `/xl/media/imageN.ext` and an image
  relationship; shapes are `xdr:sp` with `prstGeom`, solid/scheme fill or
  `noFill`, `a:ln` with width and preset dash, `txBody` with insets, anchor,
  alignment and first-run formatting; `set` rewrites only the anchor
  markers, `cNvPr` names and `xfrm` attributes in place; `remove` drops the
  anchor, the relationships only it used and the parts nothing else
  references. The worksheet's drawing part is created on demand.
- **Conditional formatting** adds a `dxf` to the styles part (the `dxfs`
  element is created at its schema position when absent).
- **Reads on both owners**: `Worksheet::comments`, `SourceWorksheet::
  {comments, page_breaks, drawing_elements}`, `Package::drawing_elements`,
  `{Workbook, SourceBackedWorkbook}::differential_formats`.
- **Fix**: an empty `<cfRule …/>` (`duplicateValues`, `uniqueValues`,
  `aboveAverage` without children) lost its `priority` and `dxfId` and
  failed to parse ("core cfRule is missing priority").
- **Prefix normalization**: the existing core CF writer and auto-filter
  writer emit `<x:… xmlns:x="…">`; the hyperlink writer then refuses the
  worksheet ("unknown direct worksheet children"). The new families write
  those owners in the worksheet root's own prefix
  (`conditional_formatting::package::replace_in_root_prefix`,
  `WorksheetLayout::in_root_prefix`); the public `replace_*` functions keep
  their exact output.
- `data_validation`: the token enums' `parse`/`as_str` are public;
  `Validation::{set_sqref, set_validation_type}`.

## 3. Unified layer: `litchi-ops`

A worksheet feature is one node kind under its sheet
(`/Sheet1/comment[N]`, `hyperlink[N]`, `picture[N]`, `shape[N]`, `cf[N]`,
`dataValidation[N]`, `autofilter`, `rowbreak[N]`, `colbreak[N]`). The node
holds one `Arc<Feature>` (kind, position, leaf value), so `Node` stays
within its 96-byte bound. Listing a kind reads only the parts it lives in
(the worksheet blocks outside `sheetData`, the comments part, the drawing
part, and the styles part's `dxfs` for a CF's formatting). Notes are
numbered row-major by cell (a comments part keys them by cell, not by
document order); everything else follows the source order. Queries scoped
to feature kinds (`query cf`, `query comment`) list them; mixed walks with
grid kinds keep the grid phases, as charts do.

Property names, defaults and aliases are OfficeCLI's (`schemas/help/xlsx`,
`ExcelHandler`): CF `type` aliases (`top10`, `topPercent`, `bottom`,
`belowAverage`, `notContains`, `dateOccurring`, …), generated text and
blank formulas, `ref`/`sqref`/`range`, `value`/`value1`/`formula1`; DV
`errorStyle` `warn`/`info`, `inCellDropdown` (inverted `showDropDown`),
list literals quoted; picture/shape geometry in cells or unit lengths.
Every table is one registry `Table`: read projection, write admission,
capabilities and the schema share it. Pending properties are declared
`Access::Gap` with `unsupported_property`, so `--unsupported warn` skips
them as OfficeCLI does: data-bar `negativeColor`, `axisColor`,
`axisPosition`, `direction` (x14), autofilter `criteria0`, picture `crop`,
`opacity`, `shadow`, `glow`, `reflection`, `softEdge`, shape
`gradientFill`, `shadow`, `glow`, `reflection`, `softEdge`.

Cell `link`/`tooltip`/`display` (registry `spec::CELL_LINK_TABLE`, write
only; the read gaps stay) and the sheet `autoFilter` (`SHEET_FILTER_TABLE`)
are staged into the same commit as the other properties of the `set`.
Unsafe link schemes (`javascript:`, `vbscript:`, `data:`, `file:`) are
`invalid_value`. Leaf refusals of authored input map to `invalid_value`,
not `corrupt_file`.

Deliberate deltas versus OfficeCLI: hyperlinks are also their own
`hyperlink[N]` element with add/set/remove (OfficeCLI: query only, written
through cell `link`); note positions are row-major; a CF `type` OfficeCLI
does not know is refused, not defaulted; a cellIs rule with no formatting
gets no `dxf` (OfficeCLI always adds one); `remove` of a CF drops only the
format, as OfficeCLI does, leaving its `dxf`.

## 4. Tests, checklist, goldens

- `tests/xlsx_features.rs` (8 tests): every kind added, read back on the
  source-backed and eager owners after a save, set, queried and removed;
  all 18 CF rule spellings; composition of every family with cell writes in
  one workbook; add-then-remove returns byte-identical bytes.
- Leaf unit and doc tests for each family (`*/edit.rs`).
- `crud_matrix`: XLSX claimed cells **92 → 191**, all ok (nine recipes).
- Goldens re-blessed, both feature sets: capabilities and registry (XLSX
  only); `officecli-capability-delta.md`; corpus outcomes and deltas.
- `pyo3-litchi`: `test_xlsx_worksheet_features`. The stubs carry no
  per-property types (`_Props` is a mapping), so no stub changed.

## 6. Measurements

`examples/xlsx_features_bench.rs`, release, `taskset -c 7`, three
processes; each process alternates leaf and unified ABBA over 20 samples
and prints medians. The workbook holds 40 notes, 40 conditional formats,
40 validations and 40 shapes. Medians (best process) and the three
unified/leaf ratios:

| scenario | leaf µs | unified µs | ratio |
|---|---:|---:|---|
| `cf_get` (cf[20] + formatting) | 754.5 | 770.7 | 1.021, 1.017, 1.018 |
| `comment_get` (comment[20]) | 38.4 | 42.6 | 1.097, 1.109, 1.104 |
| `comment_add` | 186.0 | 187.5 | 1.008, 1.005, 1.009 |
| `cf_add` (cellIs + dxf) | 2595.4 | 2600.3 | 1.002, 1.000, 1.002 |
| `link_set` (cell `link`) | 249.3 | 251.9 | 1.011, 0.990, 1.008 |
| `shape_add` | 426.8 | 431.3 | 1.010, 1.011, 1.010 |

Every write is within the 10% single-op budget. Two misses were fixed
before this table: `comment_add` was 1.40–1.46 (the adapter parsed the
comments part twice more to check for a duplicate and to number the new
note; the leaf `comments::edit::Edit::add` now refuses a duplicate and
returns the row-major position from its one parse), and `link_set` was
1.15 (the adapter read the cell's hyperlink to keep its tooltip; as
OfficeCLI, a new link now replaces the hyperlink with the command's own
tooltip/display, and only a tooltip/display-only change reads). The read
`comment_get` costs +4 µs (+10%): the adapter numbers notes row-major,
which sorts and clones the 40 notes; it is proportional to the sheet's
note count, not to the workbook. Raw lines are not committed; rerun with
the command in the example's header.

## 7. Gaps and follow-ups

1. Office 2010 conditional formatting: x14 data-bar styling
   (`negativeColor`, `axisColor`, `axisPosition`, `direction`) and x14-only
   icon sets; the core writer refuses worksheets that carry x14 CF.
2. Drawing effects, crops, transparency and gradient fills
   (`shadow`, `glow`, `reflection`, `softEdge`, `gradientFill`, `crop*`,
   `opacity`), picture `hyperlink`/`decorative` authoring, shape text
   and fill `set`, grouped shapes as `shape[N]`.
3. Threaded comments: read and write (codec only in the leaf; no package
   loader); notes are the legacy layer.
4. Note run formatting (`font.*`, `runs`) on add/set.
5. Auto-filter column criteria (`criteriaN.op`).
6. `run` (cell rich-text runs) and `sparkline`, the optional items.
7. A worksheet whose root is prefixed (`x:worksheet`): the validation
   writer emits an unprefixed owner the hyperlink writer then refuses.
8. Pivot tables, slicers and OLE stay recorded deferrals (P2).

## 8. Gates (exit 0, after merging trunk `3fa22ff1a4`)

| command | exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-xlsx -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features --all-targets -- -D warnings` (no features; `--features xlsx`) | 0, 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-xlsx -p litchi-ops` | 0 |
| `cargo test -p litchi-xlsx --all-features` | 0 |
| `cargo test -p litchi --all-features` | 0 |
| `cargo test -p litchi-ops --all-features` (`--lib`, `--doc`, then all 141 `--test` targets in chunks of 20) | 0 |
| `LITCHI_BLESS=1 cargo test -p litchi-ops --all-features --test capabilities --test registry`, and with every feature except `encryption` | 0, 0 (XLSX goldens only) |
| `LITCHI_BLESS=1 cargo test -p litchi-ops --all-features --test crud_matrix --test corpus_fidelity` | 0 (fidelity unchanged) |
| `LITCHI_BLESS=1 cargo test -p litchi-cli --all-features --test officecli_corpus`; `python3 tools/officecli_corpus.py`; `--check` | 0, 0, 0 |
| `python3 tools/officecli_capability_delta.py --officecli /home/zhuhe/code/litchi/3rdparty/OfficeCLI` | 0 |
| `python3 tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py verify` | 0, 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` (40 profiles, 238 families) | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` (280 tests) | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
