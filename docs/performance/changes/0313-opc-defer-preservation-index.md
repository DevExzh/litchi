# Change 0313: Defer OPC preservation-index validation

Status: implemented and validated

## Scope

`PreservationProvenance::from_package` no longer constructs a
`soapberry_zip::PreservationIndex` merely to validate a source archive during
ingress. It first performs an allocation-free pass over the central-directory
iterator, rejects iterator errors, counts records with checked arithmetic, and
requires the actual count to match `entries_hint()` and the supported ZIP32
entry bound. The existing fallibly reserved pass then classifies raw central
member names and preserves physical order as before.

`try_write_preserved` remains the publication boundary for physical-layout
validation. It reconstructs `PreservationIndex` from the retained source and
completes its structural checks before creating or writing preservation
actions.

## Structural effect

This removes the eager preservation-index scratch buffer and duplicate index
construction from provenance capture, reducing temporary owned-open peak
workspace and allocations. It makes no post-open residency claim: retained
provenance still owns its existing metadata and shared part baselines. Physical
layout validation is deferred until a preservation write is attempted.

No RSS, OOM, latency, or throughput claim is made here; those require a
separate benchmark or resource-profile measurement.

## Validation

- `cargo test -p litchi-opc --lib`: 251 passed.
- Strict all-feature library Clippy passed with warnings denied.
- All-feature rustdoc passed with warnings denied and without dependency
  documentation.
- Formatting and diff hygiene passed for the changed Rust files.
- All Cargo commands used `CARGO_BUILD_JOBS=1` in one isolated target and ran
  sequentially.
