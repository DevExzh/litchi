# Change 0398: unified XLSX worksheet and cell selectors

Date: 2026-09-04

Status: accepted correctness and API-unification change. This change adds no
performance claim.

`performance_claim: none`

`claim_authorized: false`

## Public API

With the `xlsx` feature enabled, the unified `litchi::sheet` facade exposes
the XLSX-gated selector:

```text
litchi::sheet::Workbook::sheet(
    name_or_zero_based_position
) -> Result<Option<SelectedWorksheet>>
```

Names use case-insensitive XLSX matching; positions are checked zero-based
workbook positions. A missing name or out-of-bounds position returns
`Ok(None)`. The returned `SelectedWorksheet` reports its canonical `name()`
and zero-based `position()`.

`SelectedWorksheet` is a lifetime-free `Clone + Send + Sync` handle over a
private `Owned`/`Source` wrapper. The eager variant bridges the existing XLSX
worksheet view into the same owned selected view; the source variant retains
the source-backed owner. Eager and source-backed workbooks therefore expose
the same selected semantics without leaking a borrowed worksheet lifetime.

## Cell and range semantics

`SelectedWorksheet::cell` accepts either an A1 reference or a raw
`(row, column)` tuple and returns an owned `SelectedCellView`. A1 is Excel's
one-based lexical notation, while raw coordinates are zero-based. The view
keeps the exact states `Missing`, `Covered(Rect)`, and `Stored(Cell)`, and the
stored `Cell` retains `Empty`, `Formula`, and `Unknown` rather than converting
through the legacy dynamic facade.

`SelectedWorksheet::cells` accepts an A1 range or raw
`(start_row, start_column, end_row, end_column)` bounds and returns an owned
sparse `Vec<SelectedCell>`. A1 endpoints are inclusive; raw bounds are
zero-based and half-open. Missing cells and merged followers are not
synthesized. Invalid coordinates and ranges remain typed XLSX errors.

Selection is catalog-only for source-backed XLSX workbooks: worksheet and
range payloads are deferred until a selected read. The selected scanner,
bounded fallback materialization, deferred-Part cache, freshness fence,
resource limits, cancellation checks, and typed source-change errors remain
owned by XLSX. A chart or other non-grid sheet returns the typed XLSX
`NotWorksheet` error. A non-XLSX runtime workbook returns the core
`Unsupported` error. The existing eager bridge has parity for the selected
view, and the legacy 1-based dynamic worksheet traits are unchanged.

## Coverage and validation

No benchmark selector is added: the selectable registry remains **419** and
the default remains **36 cases / 198 rows**. Focused current validation
includes four XLSX public tests, five public tests when XLSB is enabled, one
owned/source bridge check, one non-XLSX runtime check, and the corresponding
feature checks. The full xlsx-gated library run passed 69 of 70 tests. Its one
failure,
`source_bytes_catalog_and_selection_defer_corrupt_unselected_payload`,
reproduced identically when run alone from the exact pre-change
`c9fde90dd` baseline, so it is recorded as pre-existing rather than a 0398
regression.

The API guide contains a conceptual compile-shaped example using
`use litchi::sheet::{Workbook, SelectedCellView};` and
`workbook.worksheet_count()?`. Validation used explicit stable
rustc/Cargo/Rustdoc 1.98.1 because the pinned 1.95 toolchain lacks Cargo.

No latency, allocation, RSS, physical-I/O, cache, throughput, or
generalization claim follows. In particular, source-backed selection is not a
claim about all workbook formats, all facades, or a format-neutral worksheet
model.
