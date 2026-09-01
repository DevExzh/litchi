# Change 0364: XLSX selected-cell dependency streaming

**Date:** 2026-09-02
**Status:** Implemented
**Performance claim:** none

## Decision

Extend the selected-worksheet scan to track the maximum shared-string and
direct cell-style references across all cells, together with the target SST
index. Cold source-cell reads for plain selected SST and direct `c@s` cells
then stream the canonical `sharedStrings` and `styles` parts sequentially
through the existing OPC verified readers. These cells validate and resolve
without publishing `Store`, worksheet `PartData`, a full text `Vec`, a style
`Catalog`, or semantic dependency-cache state. Warm semantic caches no longer
rematerialize evicted `PartData`; public signatures remain unchanged.

## Processing contract

Every dependency reader reaches XML EOF and the CRC, size, source, and
cancellation fences before a value or fallback is returned. Invalid, missing,
or out-of-range references, and unsupported or oversize parts, use the
established eager diagnostics after the readers close. Rich, phonetic,
extension, and foreign SST entries, row or column styles, merges, shared,
array, and data-table formulas remain eager fallbacks.

The final cell source and cancellation fence runs even when the parser
returns an error, preserving the existing error precedence and freshness
boundary.

## Resource boundary

The dependency readers reuse the existing OPC verified-reader and eager
diagnostic boundaries. Quick-XML and current-item allocations remain bounded
only by the documented limits. This change establishes no latency, RSS, OOM,
or fixed-memory claim.

## Validation

Focused validation passed `28/28`, and library validation passed `856/856`.
Scoped Clippy passed apart from the known unrelated pre-existing `hyperlinks`
`useless_asref` issue. These are correctness and ownership observations only;
`performance_claim: none`.

## Residual scope

Only plain selected SST and direct cell-style references use the cold
dependency path. Rich or foreign shared strings, phonetic and extension
content, row or column styles, merges, shared, array, and data-table formulas,
invalid or missing references, and unsupported or oversize parts retain eager
fallback behavior. No latency, RSS, OOM, fixed-memory, or broader dependency
streaming result follows.
