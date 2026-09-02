# Change 0365: XLSX source-worksheet range streaming

- `performance_claim`: none

## Summary

Cold `SourceWorksheet::cells(area)` and staged `visit_cells(area)` now use a
verified sparse raw range scan for eligible worksheets. The scan performs its
dependency scans, reaches XML/MCE/x14ac EOF, and completes ZIP CRC/size
verification plus source/execution fences before publishing a result or
invoking callbacks. It emits sparse physical output only: missing coordinates
are omitted and explicit empty cells are retained.

The eligible cold path uses a multi-index shared-string stream and a direct
style-count stream without a worksheet `Store`, worksheet `PartData`, or
semantic dependency cache. Warm `Store` access remains the fast path.
`NotEligible` requires eager fallback only after the verified reader returns.
Merges, shared/array/data-table formulas, row/column styles, rich, phonetic,
extension, foreign, and general-reference cases remain eager. `stored_extent`
is unchanged.

`visit_cells` stages an owned `Vec`, so memory scales with selected physical
output. Its allocations are not a fixed-memory or OOM claim, and no latency or
RSS claim is made.

## Validation

Focused validation passed `27/27`; the full `litchi-xlsx` library passed
`883/883`; and package Clippy passed with `-D warnings`, with only the
unrelated `clippy::useless-asref` issue allowed. This is correctness and
boundary evidence only; `performance_claim: none`.
