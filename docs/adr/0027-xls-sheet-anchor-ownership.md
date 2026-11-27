# ADR 0027: XLS sheet-anchor ownership

- Status: Accepted
- Date: 2026-08-06

## Context

`[MS-ODRAW]` intentionally leaves `OfficeArtClientAnchor` host-defined.
For legacy XLS, `[MS-XLS]` defines that record as
`OfficeArtClientAnchorSheet`: two cell-relative endpoints plus the `fMove` and
`fSize` behavior bits. The generic OfficeArt parser already validates record
framing and retains the host record, but the XLS shape facade previously
dropped the payload and returned only shape kind, identifier, text, and group
topology.

## Decision

`litchi-xls::drawing_metadata` owns a small typed `SheetAnchor` model with
`AnchorPoint` endpoints and `AnchorBehavior`. Its codec validates the exact
`0xF010` atom identity, 18-byte payload, reserved behavior bits, BIFF8 column
range, and strict bounding-rectangle ordering. The XLS `Shape` facade exposes
the decoded value as `sheet_anchor`.

The owner is read-only and inert. It does not calculate pixel coordinates,
render shapes, execute controls, or mutate workbook streams. Encoding is
limited to one validated OfficeArt atom for snapshot/round-trip tests and
future XLS writer integration.

## Verification

Focused tests cover exact wire round-trips, reserved flags, truncated extents,
out-of-grid columns, endpoint order, and wrong OfficeArt record identity.
Package tests, formatting, diff checks, and the XLS boundary audit are the
required integration gates.
