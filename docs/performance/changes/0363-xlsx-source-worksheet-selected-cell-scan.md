# Change 0363: XLSX source-worksheet selected-cell scan

**Date:** 2026-09-02
**Status:** Implemented
**Performance claim:** none

## Decision

Route cold `SourceWorksheet::cell` through
`PartView::with_verified_decoded_reader` and the raw selected-worksheet
scanner for eligible simple scalar worksheets. Eligible cold queries do not
publish full worksheet `PartData`, `Store`, or cache state; repeated cold
queries rescan. Warm `Store` queries retain their existing fast path. Public
signatures are unchanged, including `cells`, `visit`, and `stored_extent`.

## Processing contract

Every `NotEligible` result falls back to the eager store only after the
verified reader returns and CRC, size, source, and context checks complete.
Merges, shared strings, styles, shared formulas, and rich inline values retain
their semantics through that fallback. Source, cancellation, and ZIP errors
remain primary, and the final outer fences run before a value is returned.

Zero, unrepresentable, and greater-than-2-GiB declared parts bypass the
scanner and retain the existing eager behavior, so the selected-cell path does
not introduce a lower part-size limit.

## Resource boundary

The scanner reuses the existing verified-reader and raw selected-scanner
boundaries; no dependency streaming was added. This change establishes no
latency, RSS, fixed-memory, or OOM-safety claim.

## Validation

Focused validation passed `7/7`, source validation passed `16/16`, and library
validation passed `828/828`. Scoped Clippy passed apart from the known
unrelated pre-existing `hyperlinks` `useless_asref` issue. The single-job
capped validation protocol observed no OOM; that is a protocol fact only, not
a performance or OOM-safety claim.

These are correctness and ownership observations only; `performance_claim:
none`.

## Residual scope

Cold eligible queries rescan by design, while warm store reuse remains the
existing fast path. The scanner remains limited to eligible simple scalar
worksheets; unsupported structures and ineligible declared parts retain eager
behavior. No latency, RSS, fixed-memory, OOM-safety, or dependency-streaming
result follows.
