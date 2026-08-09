# ADR 0009: ODF detection ownership and fuzz boundary

- Status: Accepted
- Date: 2026-08-03

## Context

OpenDocument detection was pre-refactor format code inside `litchi-core`. Its
optional `odf` feature enabled ZIP and XML inspection, so the neutral crate
depended directly on `soapberry-zip` and `quick-xml`. The detector's fuzz target
also lived in the core fuzz package and enabled that feature. This inverted the
dependency boundary: a concrete format implementation and its adversarial-input
coverage were owned by the common vocabulary crate.

The checked boundary ledger recorded this temporary state as migration-debt
orders 002, 003, and 010. Their exit conditions are satisfied when ODF package
and flat-XML detection, its dependencies, and its fuzz coverage all move to the
concrete `litchi-odf` crate.

## Decision

`litchi-odf-common::detect` is the sole owner of packaged and flat
OpenDocument detection. Its concise byte entry point is:

```rust
litchi_odf_common::detect::bytes(&input)
```

It returns the neutral `litchi_core::detection::FileFormat` classification, so
dependency direction remains concrete format to neutral vocabulary. ZIP and XML
parsing dependencies belong to `litchi-odf-common`; `litchi-core` no longer
declares them and no longer exposes an `odf` feature or ODF detector module.

`litchi-odf::detect` re-exports the contextual facade
`detect::{Format, mime, flat_mime, flat, bytes, reader}`. `Format` is a short
contextual re-export of the neutral classification, not a legacy compatibility
alias, so a direct `litchi-odf` consumer need not add a second dependency
merely to match detector results.
Ordinary MIME and flat classification borrow their input and do not allocate.
Packaged detection validates the required first, stored `mimetype` local entry
in place, bounds its size, checks its CRC, and validates the ZIP central
structure without copying or decompressing that payload. `reader` restores the
caller's original stream position; a restoration failure is a failed detection,
not silently changed caller state.

The future `litchi-detect` coordinator from ADR 0002 may orchestrate leaf
detectors, but it does not move ODF parsing back into core or duplicate the ODF
grammar. The umbrella uses the leaf API until that coordinator exists.

The ODF detector fuzz target moves to the existing `litchi-odf` fuzz package
and calls the canonical byte API directly. The core fuzz package remains useful
for dependency-free signature detection and no longer enables an ODF feature.
Boundary ledger entries 002, 003, and 010 are deleted. `odf` remains in the
policy's `core.format_features` denylist so CI rejects reintroducing the
resolved debt.

The legacy `litchi_core::detection::odf` path is removed without a re-export,
type alias, feature alias, or deprecated compatibility shim. This refactor is
intentionally breaking, consistent with ADR 0008.

## Consequences

- ODF callers import detection from the concrete format crate that owns the
  package and XML semantics.
- The neutral core dependency footprint and feature surface shrink.
- Common packaged detection avoids the former full `mimetype` decompression
  copy; the owned reader path still buffers a seekable stream because central
  ZIP validation requires complete bytes in the current container API.
- Fuzz compilation now verifies the public ODF owner instead of keeping the
  retired core path alive accidentally.
- Any downstream user of the legacy path must migrate explicitly; failures are
  compile-time and cannot silently select a compatibility implementation.
- Boundary-policy tests must reject a stale debt entry as well as any renewed
  core ODF feature, ZIP edge, or XML edge.

## Verification

The focused acceptance gate is the boundary checker and its unit regressions,
followed by normal checks of both fuzz manifests. Fuzz executables need only
compile during this migration slice; corpus campaigns and native Office round
trips are separate evidence and are not implied by this ownership change.
