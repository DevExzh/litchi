# Change 0307: XLSB eager worksheet catalog mapping

Status: implemented pending validation

`performance_claim: none`

## Public worksheet mapping

The eager XLSB owner keeps public worksheet ordinals separate from complete
workbook catalog positions. The complete catalog remains in workbook order and
contains every tab, while `WorkbookTrait` worksheet count, names, indexed
lookup, named lookup, and iteration expose only ordinary worksheet parts.

For a catalog such as `[Worksheet A, Chart C, Worksheet B]`, the public
worksheet sequence is `[A, B]`. Worksheet ordinal `0` resolves to catalog
position `0`, and worksheet ordinal `1` resolves to catalog position `2`.
Materialization and worksheet caches use the resolved catalog position rather
than treating the worksheet ordinal as a complete-tab position.

Structured-table ownership is normalized to public worksheet ordinals at the
facade boundary. Parsing and formula contexts remain catalog-positioned, so
source references and package metadata retain their workbook-order indexes.

## Metadata preservation

The mapping is an adapter over the public worksheet surface. It does not
renumber or discard the all-tab formula context, chart-sheet metadata, or
drawing metadata. Those owners continue to use workbook-order tab positions so
formula references and chart/drawing anchors retain their source meaning.

## Relationship-kind classification

Each workbook tab is classified from its relationship kind before the public
worksheet map is built. Worksheet relationships enter the worksheet map;
chart-sheet, dialog-sheet, macro-sheet, and international macro-sheet
relationships remain complete-catalog entries but are excluded from ordinary
worksheet selection.

## Regression scope

Generated eager facade coverage uses a mixed chart-tab workbook and checks the
worksheet-only count, names, indexed lookup, named lookup, and iterator order.
It forces the eager fallback with a Tail worksheet after a chart tab using a
sparkline operation, then asserts Tail selection and text output. The fallback
translates the worksheet handle's catalog position back to the eager owner's
worksheet ordinal before selecting the eager worksheet.
Any formula assertion in the surrounding round-trip coverage is only a smoke
check that the workbook still round-trips; it is not evidence of
catalog-sensitive current-sheet formula behavior. The regression scope makes
no additional formula-specific behavior claim.

## Measurement boundary

No performance improvement is claimed. This change corrects worksheet indexing
and preserves all-tab metadata without changing the bounded read policy.
