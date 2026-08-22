# Change 0255: source-backed PPTX owned-byte ingress

## Status

Landed in `ae6e5a7ab`. This is a correctness- and work-removal integration,
not measured before/after performance evidence.

## Scope

The ordinary `litchi::Presentation::from_bytes` and
`Presentation::from_bytes_with_limits` PPTX path now retains the caller's
owned buffer behind `OwnedSource`, opens one `SourceBackedPackage`, and hands a
`SourceBackedPresentation` to the existing unified facade variant. Opening
still reads and validates the OPC/content-type/relationship and PresentationML
catalog closure. Ordinary slide and media payloads remain unloaded until a
selected semantic query needs them.

This removes the previous mandatory `OpcPackage::from_bytes_with_limits`
fallback for valid PPTX inputs, which decompressed and retained every admitted
Part before the facade could answer a catalog-only query. The public smart
detector remains eager and source-compatible; this is a private normal-facade
handoff rather than a second opt-in API.

## Preservation and fallback gates

- Non-ZIP, non-OPC, and non-PPTX inputs drop the probe and recover the original
  `Vec` allocation for the established detector. Tests cover both an ordinary
  ZIP and a structurally valid non-PPTX OPC package, including pointer and
  capacity identity.
- A catalog identified as PPTX returns its typed PresentationML semantic-open
  failure directly instead of silently retrying through eager materialization.
- The explicit `ReadLimits` value is used by both the source-backed probe and
  the compatibility fallback. Existing OOXML-before-ODF/iWork arbitration is
  retained.
- The source-backed facade variant is target-independent for owned bytes;
  filesystem source detection remains limited to platforms with `FileSource`.
- This change does not alter save/publication behavior or claim a new
  topology-changing edit closure.

## Verification

Focused host tests covered eight source-PPTX facade cases and two allocation
recovery cases. They prove catalog queries do not load the corrupt unselected
slide, selected payloads are cached after one cold load, metadata is cached,
limits remain bounded, source revision/cancellation checks remain active, and
path/owned-source text agrees with an independently constructed eager
`OpcPackage` facade control.

The focused tests, rustfmt checks, crate-boundary gate, and `git diff --check`
passed. A narrow Clippy attempt stopped on two pre-existing
`clippy::double_must_use` errors in `litchi-opc`; no warning from this change
was reached. A cross-target build was stopped and its generated artifacts were
cleaned when disk use rose; target-independent cfg coverage was reviewed
separately.

## Claim boundary

No latency, allocation, RSS, physical-I/O, decompression-byte, or end-to-end
PPTX speedup is claimed. A future pinned release comparison must measure
catalog-only, selected-slide, full traversal, and media-heavy workloads before
quantifying the removed eager work.
