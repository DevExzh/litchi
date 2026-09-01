# Change 0362: XLSX selected-worksheet raw scan

**Date:** 2026-09-02
**Status:** Implemented
**Performance claim:** none

## Decision

Add the public narrow raw path
`litchi_xlsx::raw::selected_worksheet::{scan, ScanOutcome, SelectedCell,
NotEligibleReason, StreamResult}`. It performs one-pass active MCE+x14ac
selection through XML EOF for an eligible single-cell worksheet subset. This
is a raw capability boundary; it does not route a `SourceWorksheet` or replace
the eager worksheet parser.

## Processing contract

The scanner distinguishes a missing selected cell from an explicit empty cell,
validates strict row and cell order, and validates the supported scalar lexical
forms. x14ac `ValidateOnly` parses extension descent while avoiding a row
`BTreeMap`. Merges, styles, shared strings, shared or array formulas, rich
inline values, and unknown valid structures produce typed `NotEligible` only
after XML, MCE, and raw scanning have reached XML EOF.

`NotEligible` is not worksheet semantic validity. The caller MUST fall back to
the eager parser after that outcome; malformed XML, invalid ordering, and
invalid scalar syntax remain typed parsing or validation errors rather than a
successful selected-cell result.

## Resource boundary

The one-pass traversal is bounded by the existing XML/MCE/raw processing
limits, but it is not a fixed-memory or OOM-safe claim. quick-XML parser state,
observer allocations, and conversion allocations are outside this accounting
boundary. No source-backed `SourceWorksheet` routing, OPC verified reader,
CRC/size/source fence, or full-worksheet streaming was added.

## Validation

Focused validation passed:

- selected-worksheet focused tests: `8/8`;
- worksheet module tests: `43/43`;
- `litchi-xlsx` library tests: `821/821`.

These are correctness observations only; `performance_claim: none`. No
latency, RSS, or OOM result is claimed.

## Residual scope

The eligible single-cell subset remains deliberately narrow. Unsupported
structures return `NotEligible` and require eager-parser fallback. Full
worksheet streaming, source-worksheet routing, OPC verified-reader
integration, CRC/size/source fencing, and performance/resource measurement
remain later work.

