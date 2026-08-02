# ADR 0016: Checked BIFF8 writer locations beyond ordinary cells

- Status: Accepted
- Date: 2026-08-03

## Context

ADR 0012 made formula references grid-safe, and the first writer-coordinate
slice made ordinary cell payloads constructible only from a checked private
position. Other public writer paths still accepted wide raw row and column
values. Some validated them after borrowing or mutating worksheet state, while
data validation inserted a provisional rule and then recovered it with
`unwrap` and `pop` when later checks failed.

These paths cover different BIFF8 wire domains. An ordinary cell, merge,
validation target, data-table input, Web publication range, or RTD subscriber
uses the 65,536-by-256 cell grid. `[MS-XLS]` section 2.5.160 instead permits a
horizontal page-break column span through 16,383. Treating every location as
one generic cell coordinate would reject conforming page-break records.

## Decision

The dedicated checked values introduced or migrated in this slice carry the
narrowest representable storage for their wire domain. Their cell-grid rows
use `u16` and columns use `u8`; inclusive public ranges can be created only
through fallible constructors that reject reversal and out-of-grid values.
Horizontal and vertical page breaks use distinct checked types. Horizontal
break rows are cell-grid rows, while their column span retains the
specification's `u16` domain through 16,383. Vertical breaks use a checked
`u8` column and `u16` row span. Some established private owners retain wider
storage after a checked operation boundary and remain migration debt.

The public breaking surface follows the same rule:

- `XlsDataValidationRange` and `XlsWebPubRange` have private fields, checked
  `new` constructors, and short accessors.
- `XlsDataValidation` owns one checked `range`; `new` supplies safe message and
  alert defaults.
- `XlsDataTableInputCell::present` and `XlsRtdCell::new` narrow raw inputs
  before values enter retained state.
- Web-publication insertion is fallible for range, source, and string checks;
  real-time-data insertion validates the subscriber location and
  workbook-relative sheet in context.
- Merge, AutoFilter, filter-condition, sort, pivot, page-break, and validation
  operations validate every fallible location input before mutating a
  worksheet or its associated defined names.

The writer sorts page breaks into the order required by `[MS-XLS]` sections
2.4.142 and 2.4.343, rejects overlap, and enforces their respective 1,026 and
255 break-entry limits. Data-validation serialization consumes prevalidated
ranges; it does not repair a partially inserted rule. No compatibility aliases
or unchecked public constructors are retained.

The worksheet-view amendment applies the same policy to `[MS-XLS]` `Window2`,
`Scl`, `Pane`, and `Selection` records. The writer exposes contextual
`view::{Scale, Mode, Pane, Selection, View}` instead of public option structs
with independently mutable fields. `Scale` proves positive BIFF terms and the
10% through 400% fraction. `Pane::frozen` and `Pane::split` encode their
different cell-count and twip domains, require a real split axis, and reject an
active pane that cannot exist. `Selection` owns a nonempty bounded range list,
and `XlsSelectionRange::new` rejects reversed coordinates before retaining
them.

`View` keeps its nine display switches in a private BIFF-aligned `u16`
bitflags value. Short fluent setters preserve the concise facade; checked
setters cover origins, palette indices, zooms, panes, and selections.
`put_scale`, `put_view`, and `put_pane` move accepted values into retained state
and return the previous owned state. Pane/selection replacement validates the
complete prospective view with a non-allocating four-pane bit mask before the
first replacement. The old public option names, raw `set_zoom`,
`set_worksheet_view`, and `split_panes` entry points are removed.

## Consequences

- Dedicated checked values introduced by this slice are representable in
  BIFF8 by construction, and invalid raw locations return a typed error without
  unwinding.
- A rejected location does not change the touched worksheet collections or
  defined-name state.
- The page-break API does not incorrectly conflate a printing record's wider
  domain with the ordinary cell grid.
- Invalid view scale, pane, origin, selection grouping, active cell/range, or
  record-count state is rejected without unwinding or changing the retained
  view. Nine display options occupy the same two-byte representation written
  to `Window2`; this is a layout property, not a measured cache or latency
  result.
- Some pre-existing internals still use wide private coordinate storage after
  a checked operation boundary, including AutoFilter ranges, sort keys, pivot
  locations, and data-table anchors. Hyperlinks, row and column properties,
  conditional formatting, scenarios, consolidation, phonetic records, and
  shape anchors remain later migration work.
- `XlsSortData` still needs an explicit Rw12/Col12 policy. Data-table
  range-overlap policy, save-time RTD topic encoding failures, and unrelated
  production `unwrap` calls also remain open. Split-pane transactionality is
  closed by the checked `Pane` plus failure-atomic `put_pane` boundary.
- This is a type- and behavior-safety change. The compact fields make no
  allocation, throughput, or cache claim. Focused native open-and-inspect
  evidence applies only to the generated pane artifact described below.

## Verification

Focused adversarial tests cover exact maximum locations, overflow, reversed
and overlapping ranges, count limits, invalid workbook-relative RTD sheets,
serialization and reopen, failure atomicity, and `catch_unwind` rejection.
Selected integration suites cover data tables, validation, page setup, RTD,
Web publication, sorting, list objects, pivots, links, values, and worksheet
views. All 38 selected integration tests and the three focused unit tests pass.

The view amendment adds six unit tests, two typed writer/readback tests, five
workbook-view tests, and two lint-regression tests. They cover exact record
bytes, BIFF boundaries, compact flags, no-unwind rejection, byte-for-byte state
equality after rejected edits, move-returned old state, and frozen/split
round trips.

The `xls_styles_example` generated a BIFF8 workbook with one frozen row and
column. Through Computer Use, desktop Excel for macOS opened it in expected
Compatibility Mode without a repair prompt. After semantic navigation from the
initial used range to `M30`, row 1 and column A remained visible, confirming
that Excel interpreted the emitted frozen-pane records. No content edit or
resave was performed, and the exact application version was not recorded; this
does not certify every split-pane, selection, zoom, or view combination.

Warning-denied Clippy and rustdoc, formatting, and diff validation are green
for `litchi-xls`. Per explicit user direction, the previously green full
workspace suite is not repeated.
