# Change 0285: PPTX lazy slide catalog

Date: 2026-08-26

Status: Accepted bounded catalog behavior

Performance claim: none

## Decision

The unified presentation facade now exposes a PPTX-only `slide_catalog()`
projection. Each `SlideDescriptor` contains the zero-based presentation
position and producer-visible `p:sldId@id`. Source-backed PPTX inputs project
the identifiers from the already validated, retained lazy `SourceSlide`
handles. Eager PPTX inputs project the same values from the validated
presentation catalog.

Catalog construction does not read slide XML, DrawingML, or media payloads.
The existing `Presentation::slides()` method remains the text-bearing
materialization path and retains its prior behavior. Source-backed catalog
calls check execution and source freshness before and after projection, and
preserve typed stale-source failures plus the existing facade cancellation
error mapping at the facade boundary.
Non-PPTX presentation variants refuse the PPTX-ID catalog explicitly.

## Scope and claim boundary

This change establishes API and resource behavior only. `performance_claim` is
`none`; no latency, throughput, allocation, RSS, physical-I/O, or comparative
benchmark claim is authorized. The no-payload-read assertion is limited to the
bounded source-backed catalog operation and its existing source-cache
diagnostics.
