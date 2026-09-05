# 0425 XLSB drawing-corpus failure analysis

This is a source-only diagnosis of the failure at
`crates/litchi-xlsb/tests/workbook_structure_edit.rs:818`:

```text
source anchor count 5, expected 6
```

No Cargo command, test command, benchmark, profiler, or CPU workload was run
for this analysis. The test and the drawing-transfer implementation are
unchanged from baseline `340cc91ae2bdec338dfe7682b4b5d8c219a2d288`; the XLSB
part of change 0425 only changes fixed-width iterator representations in other
files. This failure therefore predates the migration and is an oracle/index
scope mismatch rather than evidence of an XLSB parser regression.

## Independent fixture census

The checked-in package XML and workbook relationship records give this census
of standard DrawingML anchors:

| Fixture | Catalog owner | Drawing XML anchor kinds | Raw anchors | Worksheet-transfer anchors |
| --- | --- | --- | ---: | ---: |
| `tdf108017_calcProtection.xlsb` | `ProtectedChart`, chart sheet at catalog position 1 | 1 `absoluteAnchor` | 1 | 0 |
| `universal-content.xlsb` | `Sheet1`, worksheet at catalog position 0 | 1 `twoCellAnchor` | 1 | 1 |
| `WithTextBox.xlsb` | `Sheet1`, worksheet at catalog position 0 | 1 `twoCellAnchor` | 1 | 1 |
| `testVarious.xlsb` | `mySheet1`, worksheet at catalog position 0 | 2 `twoCellAnchor` + 1 `oneCellAnchor` | 3 | 3 |
| **Total** |  |  | **6** | **5** |

The first fixture has two `BrtBundleSh` entries: `ProtectedSheet` (`rId1`, a
worksheet) followed by `ProtectedChart` (`rId2`, a chart sheet). Its
`xl/chartsheets/_rels/sheet1.bin.rels` points to `drawing1.xml`, whose one
`absoluteAnchor` is consequently catalog position 1. The other three drawing
parts are worksheet-owned and contain one, one, and three anchors respectively.
Counting opening anchor elements in the four raw `xl/drawings/drawing1.xml`
parts produces six without depending on the current parser.

## Why the existing loop reports five

The failing test loops over `0..source.worksheet_count()` and passes that same
number to `source.sheet_drawing(sheet)`. `Workbook::sheet_drawing` documents
its argument as a zero-based index in the complete workbook sheet catalog,
while `worksheet_count()` and `transfer_drawing_object` use public worksheet
ordinals. The loader deliberately stores chart-sheet drawings in
`sheet_drawings` using their full catalog position (`workbook/package.rs`
around lines 1018-1042), and stores worksheet drawings the same way (around
lines 1053-1069). The transfer API then translates its worksheet ordinal to a
catalog position before planning the source drawing.

For `tdf108017_calcProtection.xlsb`, the loop has one public worksheet ordinal
(`0`) and therefore probes catalog position `0`, where there is no drawing. It
does not reach the chart-sheet drawing at catalog position `1`. The remaining
three fixtures contribute `1 + 1 + 3` worksheet anchors, yielding the observed
five. Passing `sheet_count()` instead would incorrectly feed a chart-sheet
catalog position to a worksheet-only transfer API; it would not be a valid
oracle fix.

## Proposed strong oracle

Keep the two scopes explicit rather than lowering the assertion to five:

1. Add a corpus inventory assertion over `Workbook::sheet_drawings()` (or an
   independent raw XML census in the test helper) with per-fixture counts
   `[1, 1, 1, 3]` and total `6`. Include an explicit assertion that the first
   fixture's chart-sheet drawing has one anchor at catalog position `1`. This
   retains coverage of every parsed unique anchor, including the chart-sheet
   anchor.
2. Keep the durable transfer loop worksheet-only, with explicit per-fixture
   expected transfer counts `[0, 1, 1, 3]` totaling `5`. Resolve worksheet
   ordinals to their catalog positions before reading the inventory, or use
   fixture metadata that records the worksheet owner; do not pass a worksheet
   ordinal directly to `sheet_drawing`.
3. Continue invoking `transfer_drawing_object` and reopening the target for
   every worksheet anchor. The chart-sheet anchor should remain an inventory
   assertion, since the public transfer operation is documented and
   implemented for ordinary worksheet drawings. Supporting chart-sheet
   mutation would be a separate API scope change.

This preserves all six raw-anchor checks while accurately asserting the five
anchors covered by the worksheet-transfer contract. No production source
change is indicated by this failure.

## Test-only correction

The prepared test-only correction makes the two index spaces explicit in the
corpus table and keeps the assertions local to each fixture:

```rust
let corpus: &[(&str, &[usize], usize, usize)] = &[
    // path, worksheet ordinal -> catalog positions, raw anchors, transfers
    ("test-data/libreoffice-core/sc/qa/unit/data/xlsb/tdf108017_calcProtection.xlsb", &[0], 1, 0),
    ("test-data/ooxml/xlsb/universal-content.xlsb", &[0, 1], 1, 1),
    ("test-data/poi/test-data/spreadsheet/WithTextBox.xlsb", &[0, 1, 2], 1, 1),
    ("test-data/poi/test-data/spreadsheet/testVarious.xlsb", &[0], 3, 3),
];
```

For each row, assert that `worksheet_count()` equals the mapping length, sum
`source.sheet_drawings().iter().map(|d| d.drawing.anchors.len())`, and compare
that sum with the row's raw-anchor value before entering the transfer loop. For
the first row, additionally assert `source.chart_sheet(1).is_some()` and
`source.sheet_drawing(1).unwrap().drawing.anchors.len() == 1`; this is the
catalog-position-1 chart-sheet anchor that must not disappear from the oracle.

Then enumerate the mapping as
`for (worksheet_ordinal, &catalog_position) in worksheet_catalog_positions.iter().enumerate()`.
Look up `source.sheet_drawing(catalog_position)`, skip catalog positions with
no drawing, and call
`transfer_drawing_object(&source, worksheet_ordinal, anchor, 0)`. Record and
assert the row's transfer count, then assert the corpus totals are six parsed
anchors and five worksheet transfers. This keeps the worksheet ordinal for the
transfer API while using the full catalog position for inventory lookup.

After each transfer, replace the current existence-only check with the full
reopened inventory assertion:

```rust
let reopened_drawing = reopened
    .sheet_drawing(0)
    .expect("full reopen corpus transfer drawing");
assert_eq!(
    reopened_drawing.drawing.anchors.len(),
    1,
    "one selected corpus anchor must survive reopen for {relative} worksheet {worksheet_ordinal} anchor {anchor}",
);
```

The checked-in corpus has one transferable top-level object per selected
anchor, so this exact count retains the publication/readback check. A future
fixture containing a connector closure should carry an explicit expected
selected-graph count rather than weakening this assertion to `is_some()`. The
source and test edits are frozen for coordinator review; no verification run
is claimed here.
