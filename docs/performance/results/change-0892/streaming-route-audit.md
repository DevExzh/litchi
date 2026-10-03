# 0892 producer selected-route streaming audit

Status: read-only source audit at HEAD `0a0957ef7249b4b16ec26ed81d1fb32b7288bb0b`.
No production or harness source was changed, and no build or workload was run.
The scope is the XLSX `SourceWorksheet::visit_cells` path; iWork remains outside
the goal.

## Current fallback and first refusal

`SourceWorksheet::visit_cells` resolves a `Selection` before invoking any
callback (`crates/litchi-xlsx/src/workbook/source.rs:931-978`). A cold worksheet
uses `stream_selection` (`source.rs:1023-1068`), which calls the stoppable MCE
plus x14ac selected scanner (`crates/litchi-xlsx/src/raw/worksheet/selected.rs:282-307`).
The scanner stops at the event that records its first `NotEligibleReason`
(`selected.rs:1282-1305`). On refusal, the source and execution fences run, then
the worksheet falls through to `store` (`source.rs:1062-1068`).

The exact code-level root cause is earlier than the producer census name
suggests. `start_worksheet_child` admits only `dimension`, `sheetData`, and
merge markup; `<cols>`/`<col>` marks `Styles`, while every other worksheet child
marks `UnsupportedStructure` (`selected.rs:1388-1425`). The generated producer
part orders its children as `dimension`, `sheetViews`, `sheetFormatPr`, `cols`,
and `sheetData` (`tools/perf-baseline/src/producer_shape.rs:638-665`). MCE
semantic attributes filter the ignorable `xr`/`x14ac` attributes but leave the
ordinary SpreadsheetML `sheetViews` elements visible
(`crates/litchi-ooxml-common/src/mce/stream.rs:2959-3012`). Therefore the first
scanner mark predicted from the current source is
`NotEligible(UnsupportedStructure)` at `<sheetViews>`, before `<cols>` is
reached. The producer sidecar field
`selected_cell_stream_ineligible_on_cols` is only a byte-presence census
(`producer_shape.rs:401-436`), not an observed scanner reason; 0884 explicitly
labels the census eligibility evidence rather than a runtime profile
(`docs/performance/0884-producer-full-cell-scan-harness.md:28-32`).

If `sheetViews` is admitted, the next barriers are `sheetFormatPr`, then the
style-bearing `<cols>` branch, row `spans`, and the post-`sheetData`
`pageMargins`/`pageSetup` elements. Existing scanner behavior therefore cannot
reach an eligible full producer sheet by removing only the `<cols>` refusal.

## Work removed and semantics at risk

The verified OPC reader still drains and CRC/size-verifies the worksheet after
the stoppable selected scan (change 0658; see `selected.rs:10-23`). The fallback
then reads part data and runs x14ac capture, `process_ooxml`, and the full model
parser (`crates/litchi-xlsx/src/workbook/source.rs:1180-1213` and
`crates/litchi-xlsx/src/raw/worksheet/mod.rs:468-485`). `finish_parse` materializes
all cells and `Store::from_unsorted` sorts cells/rows and builds indexes
(`crates/litchi-xlsx/src/raw/worksheet/codec.rs:628-683`; `crates/litchi-xlsx/src/cell.rs:754-818`).
The current 0885 baseline and adopted 0891 profile are therefore valid: both
measure the actual fallback path, whose materialization is inside the timed
visit. 0891's adopted producer p50 ratio is `0.945064` (5.4936% improvement),
while its x14ac and MCE profile percentages remain diagnostic scopes
(`docs/performance/0891-x14ac-namespace-resolution.md:14-23,157-168`).

The larger work-removal candidate is a conservative producer metadata profile in
the selected scanner. It would continue the existing MCE+x14ac stream to EOF,
retain selected physical records and dependency handling, and avoid the second
full model/Store path. This is a larger potential gain than another small
namespace-lookup removal, but it still requires measurement and would retain the
selected record vector because callbacks must wait until all refusals have been
settled.

The profile must be narrow: whitelist the producer `sheetViews` subtree and
opaque `pageMargins`/`pageSetup`, validate canonical `sheetFormatPr` defaults,
admit only valid style-free/non-overlapping producer columns, and allow inert
row `spans`. Any unknown attribute/child, style inheritance, duplicate or late
metadata, ordering defect, repairable value, formula form, or dependency
failure must remain `NotEligible` and use the authoritative parser. Keep the
existing MCE/x14ac limits, verified reader, shared-string/style dependency
checks, source-version fences, cancellation checks, and callback error
precedence.

The key publication risk is tolerance reporting. Workbook-level
`SourceBackedWorkbook::tolerance_report` reports only stores already published
in each sheet's `OnceLock` (`source.rs:581-591`); it does not force a cold sheet
to parse. `SourceWorksheet::tolerance_report` does force `store`
(`source.rs:776-786`), and `Store::repair_style_indexes` records cell, row, and
column repairs (`crates/litchi-xlsx/src/cell.rs:831-891`). An eligible candidate
that silently accepts a repairable producer metadata defect would therefore
change the public report and hide a fallback-owned repair. The candidate must
only admit a proven clean metadata subset, and all future benchmark guards must
force each sheet's tolerance report outside the timer before accepting a clean
result. The existing aggregate post-timer check is insufficient for an
unpublished selected route.

## Ranked bounded qualification

1. Add a test-only runtime reason census for the generated archive; retain the
   static marker census only as package evidence. Assert the first refusal and
   the downstream refusal controls (`sheetViews`, `sheetFormatPr`, `cols`, row
   attributes, and page setup).
2. Differential the candidate against the current fallback for every producer
   sheet and the existing selected-route corpus: `cells`, `visit_cells` order,
   callback count, `4096/3686/410` value counts, digest
   `8d29b40718aa9d5b08c746f841195bba17ab8e01608b12e6093daf160c00764b`, archive
   identity, source version, and forced per-sheet tolerance reports.
3. Exercise malformed/repairable metadata, styles, ordering, shared strings,
   MCE/x14ac namespace aliases and invalid descent, worksheet relationships,
   cancellation, source mutation, XML/resource limits, and OPC CRC/size errors.
   Require identical typed errors or authoritative fallback, with zero callbacks
   before a failure.
4. Re-run the 397-worksheet/13-refusal differential before any workload. Then
   use the existing matched release protocol with producer-read, marker-free,
   ordinary XLSX, and non-XLSX controls. Candidate performance evidence is
   admissible only after the stronger per-sheet tolerance guard passes.
