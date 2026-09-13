# ODS selector search review

Baseline: `d5b5afdb4` (`feat(ods): add bounded sheet metadata transactions`).
This change optimizes selection within the existing metadata catalog. It adds
no public API, retained index, allocation, package rewrite, or serialization rule.

## Correctness argument

`collect_rows` visits recognized row containers in document order with a shared
logical-row counter. `append_row` records that counter, requires a positive row
repeat, and advances the counter with checked arithmetic. It likewise records
each cell's column start before advancing by its positive column repeat. A
published catalog therefore contains ordered, disjoint physical intervals;
repetitions do not require materializing logical cells.

The search maintains half-open index bounds and compares the coordinate with a
midpoint interval. Coordinates below its start discard the upper half; coordinates
at or beyond its exclusive end discard the lower half. A containing interval
identifies the same unique physical record as the former linear scan. Empty
collections and coordinates after the final interval take the existing fallback.
Midpoint arithmetic uses `lower + (upper - lower) / 2`.

The selected physical cell still takes precedence over implicit merge coverage.
Merge-anchor interpretation, explicit covered cells, and the fallback scan retain
their existing behavior. Worksheet names still require ambiguity checks; checked
position selectors still validate the worksheet position. Neither name lookup nor
merge fallback is claimed to be logarithmic.

Each visited interval retains the cancellation check and eight-unit Work charge.
The local ceiling and shared execution budget remain independent constraints.
Fewer comparisons can allow requests that previously exceeded a work ceiling;
the work accounting itself is not weakened.

## Verification scope

`sheet_metadata_selector_search.rs` exercises the public API independently of the
index implementation. The original metadata transaction, resource-context, and
XML-conformance targets remain the preservation and publication regressions.
The new index unit test checks below, exact, and above local Work limits for a
multi-interval search. The one-cell limit and cancellation tests remain intact.

The final command results and source hashes are recorded in
`root-verification.json`. Measurements and their workload limitations are in
[report.md](report.md). Native owner-schema and LibreOffice evidence remain with
the [metadata lifecycle batch](../ods-sheet-metadata/README.md); this lookup-only
change makes no new native compatibility claim.
