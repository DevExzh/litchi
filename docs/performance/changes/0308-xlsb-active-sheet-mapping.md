# Change 0308: XLSB active-sheet mapping

Status: implemented pending validation

`performance_claim: none`

## Active-tab parsing

`BrtBookView.itabCur` is interpreted as a zero-based position in the complete
workbook tab catalog. Parsing validates that position against the full catalog
length, including chart, dialog, macro, and international macro tabs; it is
not validated against the worksheet-only count.

When multiple `BrtBookView` records are present, the first view is the primary
active-view source for the logical workbook active tab. Later view records do
not replace that primary selection in the logical active-worksheet projection.
This change does not claim preservation of per-window view state.

The parser accepts an empty `BEGIN_BOOK_VIEWS`/`END_BOOK_VIEWS` pair and rejects
nested, unmatched, or unclosed BookViews containers. Extension/FRT records in
the container are tolerated and ignored for lossless compatibility, while
direct `BOOK_VIEW` records are parsed and validated. Every explicit `itabCur`
range is validated against the complete catalog, including the empty-catalog
case.

## Logical active worksheet

If the primary active catalog position identifies an ordinary worksheet, it is
translated through the catalog-to-worksheet map before being exposed through
worksheet APIs. The resulting active worksheet ordinal is therefore stable
when non-worksheet tabs are inserted before it.

If the primary active catalog position identifies a non-worksheet tab, typed
active-worksheet access returns the appropriate typed capability or
non-worksheet error. It does not silently select a neighboring worksheet. On
the dynamic facade, a source-freshness error takes precedence over that typed
active-tab error.

The infallible `WorkbookTrait::active_sheet_index` compatibility surface uses
the first logical worksheet ordinal as its fallback when the source active tab
cannot be represented as a worksheet index. Callers that need to distinguish a
non-worksheet active tab must use the typed active-sheet/active-worksheet
surface.

## Source and facade parity

Source-backed and eager facade paths use the same full-catalog `itabCur`
validation and catalog-to-worksheet translation. Their worksheet names,
worksheet ordinals, and active worksheet selection remain aligned for mixed
tab workbooks.

Where the writer path supplies no explicit active-sheet selection, its default
is the first worksheet. If a chart tab is first in the catalog, the first
worksheet therefore has catalog position `1`, and the emitted active view uses
`itabCur = 1`. This record makes no claim about retaining independent
per-window view selections.

## Regression scope

Generated eager, bytes-backed, and filesystem-path coverage exercises active
Tail and active chart selections in mixed chart/worksheet catalogs, together
with out-of-range active-tab rejection. It also covers the dynamic typed error
and freshness precedence. Multi-view behavior and additional BookViews grammar
negative cases remain deferred.

## Measurement boundary

No performance improvement is claimed. The change corrects active-tab
validation and worksheet-ordinal mapping without changing bounded reads or
materialization policy.
