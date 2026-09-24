# OLE1 inert streams

This batch provides bounded borrowed OLE1 reads, shared source snapshots, all five normative presentation forms, canonical authoring, and source-checked reversible edits. Linked paths and native payloads remain inert. Source admission precedes retention; authoring checks declared limits before output reservation. No-op and mutate/revert commits reuse the original source allocation.

Explicit read/preserve-only compatibility admits an embedded object with an exact empty presentation header, or one with no presentation. Neither profile enables fresh authoring of those producer deviations.

The isolated Rust 1.95.0 gate passed 522 CFB/OLE tests and 14 doctests, strict all-target/all-feature Clippy, formatting, and whitespace checks. One corpus-dependent integration test and one existing doctest were ignored. The separate v3 harness passed all four pinned native fixtures: paint and PowerPoint strictly; ReqIF only with MissingPresentation; Equation.3 only with EmptyPresentationHeader. It checks byte preservation, shared no-op source, safe inverse edits, truncation, and exact/over-limit boundaries.

The archived harness includes its code, lockfile, manifests, commands, and sanitized logs. Original and decoded native files remain in the external corpus/evidence tree; their hashes were independently verified and are archived. No timing, total-memory, rendering, or universal producer compatibility claim is made. Earlier evidence versions remain untouched.
