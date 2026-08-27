# Change 0306: XLSB mixed-tab worksheet mapping

Status: implemented pending validation

`performance_claim: none`

## Contract

The source-backed XLSB catalog has two deliberate coordinate spaces:

- Full catalog positions enumerate every workbook tab in workbook order,
  including chart, dialog, macro, and international macro sheets.
- Worksheet ordinals enumerate only ordinary worksheet parts, in the same
  relative order as their full catalog positions.

The public worksheet contract is worksheet-only. `worksheet_count`,
`worksheet_names`, `worksheet_by_index`, `worksheet_by_name`, and worksheet
iteration must never expose a non-worksheet tab. The all-tab `sheet_*` APIs
remain available for callers that need to inspect the complete workbook tab
catalog.

## Mapping invariant

For every worksheet ordinal `i`, the adapter resolves the corresponding full
catalog position through the source owner's worksheet-position map before
materializing or invoking the eager compatibility path. A workbook such as
`[Worksheet A, Chart C, Worksheet B]` therefore has worksheet names
`[A, B]`, with worksheet ordinals `0 -> A` and `1 -> B`, while full catalog
positions remain `0 -> A`, `1 -> C`, and `2 -> B`.

Names and cached worksheet handles use the same mapping. This prevents a
non-worksheet tab from shifting indexed lookup, named lookup, cache slots, or
eager fallback selection.

## Source text behavior

The source-backed `text` and `write_text_to` projections visit only the
worksheet-position map. Non-worksheet tabs are retained in the catalog but
are skipped by ordinary worksheet text extraction; they must not cause text
conversion to fail merely because they are present.

## Regression coverage

Generated chart-tab facade coverage checks worksheet-only count, names, indexed
and named selection, iterator order, and exact source-backed text output.

Direct correction of the eager owner `WorkbookTrait` implementation and
fallback-specific regression coverage are deferred to a follow-up change. This
record does not claim that the eager owner's internal all-tab formula context
has been changed.

## Measurement boundary

No performance improvement is claimed. The mapping is a correctness and
compatibility fix only; the source-backed bounded-read and cache policies are
unchanged.
