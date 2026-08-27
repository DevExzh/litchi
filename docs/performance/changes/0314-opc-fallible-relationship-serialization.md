# Change 0314: OPC fallible relationship serialization

Status: implemented and validated

`performance_claim: none`

## Fallible serialization boundary

The OPC relationship serializer now has a crate-private byte path that uses
fallible reservations for its sorted relationship references and serialized
output. It emits XML literals and the five standard XML entities directly,
without temporary escaped `String` values or infallible `String` growth. The
existing public `Relationships::to_xml` API remains unchanged for compatibility.

Eager package publication and source-backed relationship topology publication
use the fallible byte path before any output reaches their sink. Canonical
relationship provenance uses the same path and retains the resulting `Vec<u8>`
directly, avoiding an infallible `into_boxed_slice` shrink allocation.

## Preserved behavior

Relationship order remains deterministic by `rId`; headers, attribute order,
escaping, target-mode spelling, and empty relationship output remain byte-for-
byte compatible with the prior canonical serializer. An allocation failure in
eager publication or source-backed planning propagates through the existing
typed `OpcError::Allocation` path before publication begins.

Provenance construction remains best-effort. If canonical relationship bytes
cannot be built, the whole targeted-preservation provenance is disabled while
the exact source authorization remains intact. Later mutation therefore keeps
the existing preservation refusal/fallback semantics and never exposes a
partially constructed provenance graph.

## Measurement boundary

No throughput, RSS, peak-memory, OOM, allocation-count, or speedup claim is
made. Validation is limited to serializer parity, deterministic insertion
ordering, empty output, explicit-empty relationship-member behavior, and the
existing package/source-backed publication contracts.

## Validation evidence

- `cargo test -p litchi-opc --lib`: 255 passed.
- Strict all-feature library Clippy completed with warnings denied.
- All-feature no-deps rustdoc completed with warnings denied.
- Rustfmt and diff hygiene checks passed.
- Cargo work was run sequentially with `CARGO_BUILD_JOBS=1`, using one
  isolated target directory per run.
