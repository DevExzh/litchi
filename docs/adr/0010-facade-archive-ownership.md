# ADR 0010: Archive ownership below the facade

- Status: Accepted
- Date: 2026-08-03

## Context

The `litchi` facade directly used `soapberry-zip` while coordinating OOXML,
OpenDocument, and iWork detection. Its `ooxml`, `odf`, and `iwa` features all
activated that implementation dependency, and some probes shared one concrete
archive reader across format branches. Boundary-debt order 048 recorded this
temporary inversion.

A facade may coordinate format capabilities and return one concise result, but
it must not own their ZIP grammar or expose a container implementation type.
ODF and iWork now provide format-owned byte and reader probes. Existing OOXML
probing already opens an `OpcPackage` through the OOXML package API and inspects
its validated content types; it does not need raw ZIP traversal in the facade.

## Decision

Smart detection delegates complete input to the enabled concrete owner:

- `litchi-odf` owns packaged and flat OpenDocument detection;
- `litchi-iwa` owns iWork archive and application detection;
- the existing OOXML package API owns OPC decoding, while the coordinator maps
  validated main-part content types to the neutral format classification.

The typed iWork leaf facade is
`litchi_iwa::detect::{Format, bytes, reader, path}`. `Format` is the contextual
enum `Pages | Keynote | Numbers`, not a generic numeric discriminator or a raw
archive marker. The reader probe starts at byte zero and restores the caller's
original position; inability to restore it yields non-detection rather than a
successful result with surprising cursor state.

iWork package detection validates the root `DocumentArchive` envelope rather
than treating `CalculationEngine.iwa` as a Numbers discriminator. Pages and
Keynote can legitimately own calculation-engine components for embedded
tables, so filename precedence alone is neither safe nor complete. Versioned
component names and pre-iWork '13 nested `Index.zip` packages remain leaf-crate
concerns and are covered by the same bounded detector.

The facade removes `soapberry-zip` from its manifest and feature definitions.
It does not replace that edge with `litchi -> litchi-opc`, nor does it re-export
an archive reader as a compatibility seam. While the OOXML migration host still
exists, its already-recorded order-047 edge can supply the package API; this
decision neither makes that host canonical nor expands its debt. A future
`litchi-detect` coordinator may replace facade-local orchestration, but leaf
formats continue to own their container semantics.

Tests may build fixtures through an existing format/package writer or an
external development-only ZIP tool. Test convenience does not justify a normal
facade dependency on a concrete archive implementation.

Boundary-debt order 048 is deleted rather than converted into a canonical edge
or retained as a compatibility alias.

## Consequences

- Enabling OOXML, ODF, or iWork no longer gives the facade a direct archive
  implementation dependency. Their owner crates may still depend on ZIP
  capabilities internally.
- Format-specific validation, limits, and malformed-input policy remain with
  the format that can interpret them.
- The smart-detection API can remain concise while its implementation composes
  leaf results instead of traversing archive entries itself.
- An input that fails earlier candidates may require more than one container
  scan. The former shared reader could avoid some repeated central-directory
  work, so this change makes no latency, allocation, or throughput
  improvement claim.
- Representative mixed-format benchmarks and profiles must determine whether
  repeated probing is material. If it is, optimization belongs in a focused,
  opaque detection plan or neutral container capability; raw archive types and
  a facade implementation edge must not return merely to reduce an unmeasured
  cost.

## Verification

The dependency checker must report the removed edge as resolved only after the
manifest no longer declares it, then pass after order 048 is deleted. Format
detection tests separately establish semantic parity and cursor restoration;
the boundary change alone is not evidence of native-format compatibility or a
performance improvement.

## 2026-08-08 amendment: focused iWork detector

The iWork detector named above has moved out of the migration host. The current
leaf is `litchi-iwa-detect::{Format, bytes, reader, path}`; statements in this
record naming `litchi_iwa::detect` describe the historical ownership at the
time of the decision. Shared physical ZIP and IWA limits live below the three
concrete format owners, while semantic malformed-input policy remains with
Pages, Numbers, or Keynote. The root facade composes the focused detector and
does not regain a raw archive dependency.
