# 0362ul94c — Custom XML producer/error audit

This source-only follow-up is based on the combined OOXML publication tip
`0db7ff3a6aee574102ae0319da2c1d9e14e5e4f2` (tree
`249c3b152e87c2ff4c71fd2ee86d58cd55488e79`). It closes the remaining
Custom XML producer boundary identified by 0362ul94b.

## Classification closure

* `Edit` target misses and occupied data, properties, relationship, and part
  destinations now use content-free `Refused` values with `NotFound` or
  `UnsupportedTarget`. A package-root destination is rejected at the same
  semantic boundary.
* A typed properties schema edit refuses `UnsupportedContent` when the source
  properties part contains opaque markup that the deterministic writer cannot
  preserve. An unchanged schema topology still uses the bounded itemID splice,
  retaining every source byte; limits and XML decoding failures are unchanged.
* A shared data part that would require two different staged payloads is an
  `UnsupportedContent` refusal. Malformed relationship graphs, aliases,
  duplicate IDs, malformed XML, and invalid content types remain their
  corruption/value variants rather than being hidden as capability refusals.
* Missing captured host parts and changed captured host bytes are
  `PatchConflict` during replay. `Snapshot::load_scoped` suppresses only an
  expected `PartNotFound` hint; deferred decode and other OPC failures still
  propagate. Prepared relationship/materialization and projection mismatches
  are `Internal` invariants.

## Regression coverage

The Custom XML tests cover typed missing/occupied target refusals, an occupied
host relationship, opaque properties schema editing, shared data-part refusal,
missing and changed captured hosts during patch replay, generated projection
`Internal`, malformed relationship preservation, exact source bytes, and
failure atomicity. Existing signed, inverse-patch, opaque-payload, limit, and
malformed XML tests remain in place.

Only source inspection, `rustfmt --edition 2024`, and `git diff --check` were
run for this record. No Cargo, rustdoc, glossary, native, or performance
command was run; composed validation remains the coordinator's responsibility.
