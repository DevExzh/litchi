# 0316 — XLS and ODP structural follow-ups (record 0308)

- Status: implemented on `uo2/xls-odp-gaps`, cut from `feat/unified-ops-api`
  at `30eed270b5`; trunk merged at `5a118f742e` (`cbbdcd1fde`, record 0314).
- Date: 2026-10-16
- Input: record [0308](0308-office-structural-gaps.md) §8 follow-ups 1
  (XLS `AnchorBehavior` names), 2 (XLS shifts and reader strictness) and 4
  (ODP formatting).
- Owned: `litchi-xls` (read and write), `litchi-odp`, and the XLS and ODP
  write paths of `litchi-ops` (`format/odp*`; no XLS adapter source
  changed); the XLS and ODP rows of the checklist and fidelity goldens.
- `performance_claim`: none; §6 shows the touched paths.

## 1. Result

| # | item | outcome |
|---|---|---|
| 1 | XLS `AnchorBehavior` names | renamed after MS-XLS/POI/ECMA-376: `MoveAndResize` (0), `MoveWithoutResize` (2), `Absolute` (3); the writer's duplicate `writer::shape::Behavior` is deleted; old names banned (§2) |
| 2 | XLS shifts on producer workbooks | CF/CF12/CFEx and DV formulas, Array and Table records, position-locked anchors over rows of one height; row add+remove byte-exact **227 → 254** of 305 sheets, columns **227 → 253** (§3) |
| 3 | XLS reader strictness | relaxed with typed repairs where the spec allows (WINDOW2 zoom, stale `fShrFmla`, `fApplyAutoFilter`); kept and documented where it does not (CONDFMT count, out-of-place ShrFmla); embedded chart caches no longer read as cells (§4) |
| 4 | ODP formatting | shapes, table cells and notes take and read formatting as a whole; link and field runs take formatting; reads resolve the style chain through common styles (§5) |

## 2. `AnchorBehavior` (F1)

`[MS-XLS]` 2.5.193 sets `fMove` (bit 0) and `fSize` (bit 1) when the shape
is *kept intact* as its cells move or resize, and `fSize` MUST be set when
`fMove` is. The public read enum `drawing_metadata::AnchorBehavior` named
wire 0 `Fixed` and wire 3 `MoveAndSize`, the inversion record 0308 fixed in
the shift. It is now:

| wire | name | Apache POI `AnchorType` | ECMA-376 `ST_EditAs` |
|---|---|---|---|
| 0 | `MoveAndResize` | `MOVE_AND_RESIZE` | `twoCell` |
| 2 | `MoveWithoutResize` | `MOVE_DONT_RESIZE` | `oneCell` |
| 3 | `Absolute` | `DONT_MOVE_AND_RESIZE` | `absolute` |

The writer carried a second enum, `writer::shape::Behavior`, with the same
inversion; it is deleted and `writer::shape::Anchor::{new, cells,
behavior}` take and return `AnchorBehavior`. Every caller keeps its wire
value (the writer default and the examples write wire 3, now spelled
`Absolute`), so written bytes are unchanged. Per ADR 0034 the old names are
deleted, not aliased: `xls-minimal` bans `Fixed`, `Size` and `MoveAndSize`
under both `AnchorBehavior` paths and `writer::shape::Behavior`; the
glossary section "XLS anchor behaviour (record 0316)" lists the rows.

## 3. XLS row and column shifts (`cell_values/structural/{rules,sizes}.rs`)

| construct | before (0308) | now |
|---|---|---|
| `CF` formulas with cell references | refused | a `PtgRefN`/`PtgAreaN` component keeps its offset when every cell of the CONDFMT range and the cell it names lie on one side of the shift point (0308's shared-formula rule, now shared code); absolute references shift; a `PtgRef` with relative flags (`RgceLoc` stores coordinates, LibreOffice reads CF `PtgRef` relative parts as offsets) is admitted only when no reading reaches the shift point |
| `DV` formulas with cell references | refused | the same over the bounding box of the DV's `sqref` |
| `CondFmt12`/`CF12` | refused (in any sheet when reachable) | the FRT `ref8`, `refBound` and `sqref` shift; `rgce1`, `rgce2`, `fmlaActive` and every `CFVO` formula of a colour scale, data bar or icon set are located byte for byte (a layout the walk does not account for is refused) and shift as CF formulas |
| `CFEx` | its `ref8` was not shifted | shifts with its CondFmt range |
| CF, CF12, DV, ShrFmla, Array in another sheet | refused when an XTI reaches the shifted tab | their 3-D references into the shifted tab shift (absolute components; a relative 3-D component is refused); their same-sheet references are their own sheet's |
| `Array` | refused | the range moves whole with its row (a shift point inside it is refused); its tokens shift like a cell formula's; a non-empty `rgcb` is refused |
| `Table` | refused | the range moves whole with its formula row and column; the input cells shift as points unless marked deleted; the second input exists only with `fTbl2` |
| `PtgExp`/`PtgTbl` anchors | patched after the row table was regenerated | moved with their row, so the row-block planner sees an Array or Table next to its anchor |
| position-locked anchor (`Absolute`) | refused unless the shift is after it | keeps its cell coordinates when every row (column) from the shift point to its last row (column) has the same size after the shift as before (`ROW`/`DefaultRowHeight`, shifted `COLINFO`); a batch with several shifts or with layout changes still refuses it |
| `MoveWithoutResize` spanning the point | refused | its far end keeps its place under the same size check |
| spec-invalid `fMove` alone (1) | refused | read as POI's `DONT_MOVE_DO_RESIZE` and LibreOffice: the top-left keeps its place, the bottom-right moves with its cell |
| empty sheet with a lone `DBCELL` | `corrupt_file` (row-block order) | keeps its row table and INDEX while it stays empty |
| inclusive DIMENSIONS (last row and column, not one past) | not byte-exact | the convention is kept while it still describes the cells |
| producer INDEX `ibXF` that does not locate DEFCOLWIDTH | rewritten | kept |

Census (`CENSUS_FORMAT=xls cargo test -p litchi-ops --all-features --test
structure_census -- --ignored --nocapture`), 305 sheets, add then remove at
the start:

| | before (0308) | after |
|---|---:|---:|
| row add+remove byte-exact | 227 | **254** |
| row add+remove not byte-exact | 2 | 5 |
| column add+remove byte-exact | 227 | **253** |
| row remove existing (empty first row) | 5 | 6 |

The three new value-exact sheets are `SimpleWithColours.xls` (three
copies), newly read (§4): its producer wrote negative `DBCELL` row offsets
that a regenerated row table does not reproduce. Remaining row refusals:
13 formulas the token reader cannot read (11 `Lbl`/cell formulas, nine of
them the three copies of `ConditionalFormattingSamples.xls`; two rule
formulas of `50939.xls`), 6 non-comment `Obj` records (AutoFilter and
form-control drop-downs), 3 `RgbExtra` formulas, 6 sheets the reader
still drops (§4), 2 position-locked anchors over rows of different
heights, 2 Sort (`0x0090`) and 1 Pane (`0x0041`) dependencies, 2 3-D
references spanning sheets, 2 internal hyperlinks, 1 split OfficeArt
record, 4 `corrupt_file` fonts, 4 tabs without a grid.

## 4. XLS reader strictness, re-checked

| check (0308 kept) | spec | now |
|---|---|---|
| WINDOW2 zoom outside `0`/`10..=400` (`SimpleWithColours.xls` ×3, `SingleLetterRanges.xls`: page-break zoom 1) | "MUST be 0 or 10–400", but product notes 144/145: Excel 97–2010 truncate both fields into a byte; the page-break zoom is undefined and MUST be ignored under page-break preview with SCL; LibreOffice and POI read any value | **relaxed**: reads as the default zoom, recorded as `WorksheetDefect::ZoomLevel` (ordinal: the record's stream offset, observed: the raw value); record bytes kept |
| `fShrFmla` without a leading `PtgExp` (`ConditionalFormattingSamples.xls` ×3 tabs, `WithTwoCharts.xls`, `DateTimeToNumberTestCases.xls`) | "If this formula is part of a shared formula, rgce MUST begin with PtgExp" — a formula whose `rgce` is complete is not part of one; POI's `FormulaRecordAggregate` documents Excel writing the stale flag and clears it | **relaxed** for every profile: the flag is ignored (the formula reads as ordinary), recorded as `WorksheetDefect::SharedFormulaFlag` (row, column); the explicit `SharedFormulaFlagWithoutPtgExpV1` profile still keeps the flag and reports `formula_metadata::Defect` |
| Feature11 `fApplyAutoFilter` without `fAutoFilter` (`ConditionalFormattingSamples.xls` tab 7) | only `fPersistAutoFilter` requires `fAutoFilter`; note 185: Excel 2007/2010 always write `fApplyAutoFilter` = 1 | **relaxed**: not a defect, so nothing to record |
| CONDFMT `ccf` above 3 (`formats.xls`, LibreOffice-written, `ccf` = 4 with four CF records) | "MUST be ≥ 1 and ≤ 3", unconditional, no product note | **kept**: the count delimits the CF collection (ABNF); a reader would have to read past a MUST the spec gives no reading for |
| ShrFmla not after its anchor Formula (`15228.xls` tab 8) | "preceded by a single Formula record that specifies the first cell in the range" | **kept**: the record's range starts two columns before the preceding formula and overlaps another ShrFmla; which shared formula those `PtgExp` cells use is ambiguous |

Embedded chart substreams (`BOF`…`EOF` after a drawing object) are now
skipped by the source-backed reader's scans and by the `cell_values`
scanner, as the eager parser has skipped them since record 0293. Their
chart-cache `Number`/`Label` records were read as cells: the source-backed
`WithFormattedGraphTitle.xls` read `20` for A1 (the chart's first value),
which holds the shared string `Some Text`, and the editor refused the
sheet for "duplicate Number records". 2,004 such records sit in six corpus
workbooks (1,926 in `external_name.xls`).

## 5. ODP formatting (`litchi-odp::scene`, `litchi-ops` ODP)

- **Common styles.** `CommonStyles::parse` indexes `office:styles` of
  `styles.xml` (or a flat document), with each family's
  `style:default-style` under the empty name. `AutomaticStyles::resolve`
  walks `style:parent-style-name` (an automatic style's parent is looked up
  among automatic styles, then common ones; a common style's among common
  ones), then the family's default style; a missing parent ends the chain
  and a cycle is cut after 32 styles. `resolve_paragraph` and
  `resolve_run` apply it to a paragraph and to a run's own style (a span, a
  formatted field, or a link's text span over its link style).
  `TextProperties::inherit` fills unset properties; the font
  (`style:font-name`/`fo:font-family`) is one property. Values are
  `Arc<str>`, so resolution copies no text.
- **Reads.** Paragraph reads resolve their chain; run reads resolve the
  run's chain over its paragraph's (record 0308 read local properties
  only). Shape, placeholder, table cell (`tc`) and notes nodes gain `bold`,
  `italic`, `underline`, `size`, `color`, `font`: the value every run of
  their text shares (a paragraph without runs counts by its own), `None`
  when they differ; it is resolved once per node, on the first such key.
  A shape's graphic or presentation style is not part of the chain (§7).
- **Writes.** A set on a shape, cell or notes formats every paragraph (a
  derived paragraph style) and every run whose own chain sets a different
  value for a property the set changes (so each run reads the new value);
  a run that already agrees, or inherits, is left alone.
- **Links.** `Change::SetTextProperties` on a link keeps the link's own
  (link character) style; the one `text:span` its content is carries the
  formatting (derived), or a new span is put inside the `text:a`. The scan
  records that span (`Run::link_text_style_name`).
- **Fields.** A bare field run is wrapped in a new span. A `text:span`
  whose whole content is one field with text now reads as that field
  (`RunKind::Field`, its style the span's), so text sets stay refused and
  `Change::SetFieldText` (field refresh) rewrites the field inside its span.
  No corpus content page holds such a span (two flat masters do); query
  digests are unchanged.

Tests: `litchi-odp` `scene::tests::{link_and_field_runs_take_formatting,
formatting_resolves_through_common_styles}`; `litchi-ops`
`tests/odp_character_format.rs::{shapes_cells_and_notes_take_formatting_as_a_whole,
links_fields_and_common_styles}` (a flat document with a common parent
style, a link and a field; inverse restores the bytes; no-op; cell; notes).

## 6. Measurements

`examples/xls_odp_followups_bench.rs` (compiles against the base tree too)
and `examples/xls_odp_followups_new_bench.rs`; release, `taskset -c 9` on a
shared 32-core host (load 3–6), 15 rounds after 3 warm-ups, medians of
per-call means, paired legs ABBA, binaries alternated B A A B B A. "Before"
is base `30eed270b5` with the example copied in. Evidence, binary hashes and
the build command: [results/0316/summary.txt](results/0316/summary.txt)
(`r3-*` are the final runs, on the merged tree).

Existing paths, after/before (three runs each):

| scenario | before µs | after µs | after/before | unified/leaf after |
|---|---:|---:|---:|---:|
| XLS complete open, 20k cells | 5,416 | 5,504 | 1.016 | — |
| XLS complete open, FormulaEvalTestData | 1,786 | 1,791 | 1.003 | — |
| XLS source-backed walk, 20k cells | 895 | 904 | 1.010 | — |
| XLS row insert, 10k cells (leaf / unified) | 10,076 / 10,084 | 10,165 / 10,140 | 1.009 / 1.005 | 0.997–1.002 |
| XLS row insert, `SimpleWithImages.xls` (leaf / unified) | 146.7 / 147.7 | 146.5 / 147.5 | 0.998 / 0.999 | 1.006–1.010 |
| ODP `Scene::parse`, 200 slides, 400 styles | 495.5 | 487.0 | 0.983 | — |
| ODP `bold` of 200 runs (unified, resolved) | 237.3 | 228.8 | 0.964 | — |
| ODP bold + size on one run (leaf / unified) | 1,736 / 1,740 | 1,738 / 1,742 | 1.001 / 1.001 | 1.001–1.003 |

New paths, unified against leaf (two runs):

| scenario | leaf µs | unified µs | ratio |
|---|---:|---:|---:|
| XLS row insert, 10k cells + array formula + data table | 10,551 / 10,587 | 10,536 / 10,531 | 0.999 / 0.995 |
| XLS row insert, `WithConditionalFormatting.xls` (CF, CFEx) | 226.5 / 227.0 | 228.4 / 228.4 | 1.008 / 1.006 |
| XLS row insert, `validation.xls` (DV `PtgRefN`) | 176.8 / 176.6 | 177.8 / 178.7 | 1.006 / 1.011 |
| XLS row insert, `NewStyleConditionalFormattings.xls` (CF12) | 564.3 / 559.7 | 567.4 / 563.2 | 1.005 / 1.006 |
| XLS row insert, `SimpleWithComments.xls` (position-locked) | 112.7 / 112.9 | 114.4 / 113.9 | 1.016 / 1.009 |
| ODP bold on a whole shape, 200 slides | 1,950 / 1,958 | 1,955 / 1,957 | 1.002 / 1.000 |
| ODP open + shared `bold` of 200 shapes | 1,086 / 1,091 | 1,181 / 1,185 | 1.088 / 1.086 |

- **First version misses, fixed.** The first build (r1) cost the 10k-cell
  row insert +2.6–3.1% (the regenerated row table was parsed again to
  detect an empty sheet, the top-level records scanned for cells, and every
  record checked against the rule kinds), the 20k-cell open +3.0% (the
  WINDOW2 zoom defects drained after every record) and the 200-run ODP read
  +33.7% (style chains cloned their names and values and looked each style
  up three times). Each is now done only when it can matter, and style
  values are shared `Arc<str>`: +0.5–0.9%, +1.6% and −3.6%.
- **XLS complete open +1.6%** (5,416 → 5,504 µs, three runs each) is
  reported, not claimed as noise: the worksheet parse now decodes each
  Formula with its flag defect and tests each record for WINDOW2.
- **ODP shape read 1.09** is a whole open-and-query read (the leaf leg
  parses the scene and styles and resolves one property per shape; the
  unified leg detects the format, builds the handle and every shape's full
  `format` projection, all six character keys from one resolution). It is
  under the 1.10 review trigger record 0262 uses for full queries; no
  budget is claimed.

## 7. Gaps and follow-ups

1. **XLS:** `Lbl`/cell formulas the token reader cannot read (11 sheets,
   mostly `ConditionalFormattingSamples.xls`); non-comment `Obj` records
   (AutoFilter and form-control drop-downs); `RgbExtra` formulas; Sort,
   Pane and AutoFilter dependencies; position-locked anchors over rows or
   columns of different sizes (needs a size-to-position conversion that
   rounds); `rgcb` of array formulas; CONDFMT `ccf` > 3 and the
   out-of-place ShrFmla (spec-strict, an owner decision); producer
   `DBCELL` offsets (`SimpleWithColours.xls`, value-exact).
2. **ODP:** text inherited from a shape's graphic or presentation style
   (`draw:style-name`, `presentation:style-name`) and from
   `draw:text-style-name`; the `style:font-name` font-face to family name.

## 8. Gates (exit 0)

Run on the merged tree (`5a118f742e` and later); cargo under the
memory-capped scope shim.

- `cargo fmt --all --check`.
- `cargo clippy -p litchi-xls -p litchi-odp -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings`.
- `cargo clippy -p litchi-ops --no-default-features [--features F] --all-targets -- -D warnings` for no features and F = `xls`, `odp`, `xls,odp`.
- `cargo test --all-features` of `litchi-xls`, `litchi-odp` and `litchi`; `litchi-ops` `--lib` (922), `--doc` (210) and its 139 integration tests in seven groups of 20 (1,263 tests).
- `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-xls -p litchi-odp -p litchi-ops`.
- `tools/check_crate_boundaries.py`; `tools/non_iwork_gate.py verify`; `RUSTC_BOOTSTRAP=1 tools/check_api_glossary.py` (40 profiles, 2,313 bans, 228 families); `tools/test_no_3rdparty_references.py`; `tools/unified_bench_matrix.py --check`.
- `crates/pyo3-litchi/tests/run_tests.sh` (279 tests).
- `cd tools/perf-baseline && cargo check --locked --all-targets`.
- Goldens re-blessed after review: capabilities and registry (ODP; both feature sets), `officecli-capability-delta.md` (ODP), `results/0289/crud-checklist.tsv` (ODP `tc` set), `results/0289/fidelity.{md,tsv}` (XLS structure byte-exact 80 → 92, value-exact 2 → 5, refused 37 → 22; revert byte-exact 103 → 112, refused 13 → 4; no file got a worse outcome), `tests/golden/query_corpus/digests.json` (94 XLS digests, cells 118,712 → 123,192: newly read worksheets, less 2,004 chart-cache records formerly read as cells; ODP unchanged).
- `cargo test -p litchi-cli --all-features --test officecli_corpus` and `tools/officecli_corpus.py --check` unchanged.
