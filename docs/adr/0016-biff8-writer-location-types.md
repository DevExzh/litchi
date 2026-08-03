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

The shape-anchor amendment replaces the public field bags with
`writer::shape::{Point, Anchor, Behavior, Rect}`. `Point` proves the BIFF8 cell
and offset domains in a six-byte representation; `Anchor` proves strict
top-left/bottom-right ordering on both axes. `Behavior` exposes only the three
OfficeArt flag combinations with defined semantics, and `Rect` rejects a
degenerate group coordinate space. Shape, group, and comment insertion reserve
all retained collections and strings before mutation or object-ID allocation.
Requested child IDs are also reserved before automatic group IDs are chosen,
so insertion order cannot create a later collision.

The SortData amendment gives its packed signed four-byte field private checked
`Rw12` and `Col12` components. Public `writer::sort::{Row, Col, Range, Axis,
Method, Parent, Dxf, IconSet, Icon, On, Key, Config}` keeps those wire details
out of the facade. `Range::new` accepts inclusive checked row and column ranges;
`Key::col` is valid for row sorting and `Key::row` for column sorting, making an
ambiguous two-dimensional key unrepresentable. `put_sort` consumes a complete
configuration and returns the previous one, `remove_sort` is idempotent, and
`sort` lends the current state. The older root `XlsSort*`, `set_sort_data`, and
`sort_data` surface is removed; the separate legacy `set_sort` BIFF record is
not a SortData alias and remains intentionally distinct.

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
  a checked operation boundary, including AutoFilter ranges, pivot locations,
  and data-table anchors. Hyperlinks, row and column properties, conditional
  formatting, scenarios, consolidation, and phonetic records remain later
  migration work.
- Shape/group/comment and SortData collection growth is failure-atomic, but an
  ordinary `Write` implementation can still fail after emitting a prefix of a
  BIFF stream. Data-table range-overlap policy, save-time RTD topic encoding
  failures, and unrelated production `unwrap` calls remain open. Split-pane
  transactionality is closed by checked `Pane` plus `put_pane`.
- This is a type- and behavior-safety change. The compact fields make no
  allocation, throughput, or cache claim. Focused native open-and-inspect
  evidence applies only to the generated pane and shape artifacts described
  below.

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

The shape and SortData amendment adds 39 focused tests: six SortData unit
tests, three shape-anchor units, six shape-group units, three SortData
integrations, seven list-object integrations, five shape-writer integrations,
four group-writer integrations, and five comment-writer integrations. They
cover exact packed bounds, axis/key coupling, wire round trips, invalid flag and
rectangle states, allocation failure ordering, explicit/automatic ID
collisions, and move-returned CRUD state. The complete `litchi-xls` target also
builds without running unrelated suites.

The `odraw_native_smoke` example generated `odraw-smoke.xls` with a rectangle,
text box, and grouped ellipse/text pair through the checked anchor API. Through
the Computer Use skill, desktop Microsoft Excel for macOS opened it without a
repair prompt in expected Compatibility Mode and rendered all objects, fills,
text, and group placement. This is open-and-inspect evidence for that artifact
only: Excel edit/resave, Litchi reverse-read after resave, an Office-version
matrix, SortData UI behavior, and performance were not tested.

Warning-denied Clippy and rustdoc, formatting, and diff validation are green
for `litchi-xls`. Per explicit user direction, no redundant manual full-
workspace run was scheduled; the repository's mandatory pre-commit hook later
ran and passed the workspace lib/integration and doctest gates.
