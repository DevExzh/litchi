# 0362uj — Consumer-v8 compiler repairs

This source-only repair starts at integration commit
`7abf64f495c298cf5591ab1f73ef3ed4cd5472b2`. The v8 workspace diagnostic
reported consumer migrations that had not followed the typed ODF error
boundary: the formula and ODB leaves still sent common-core edit errors
through the ODF package mapper. These call sites now use the core staging
mapper, while package-returning ODF errors keep the typed ODF mapper. The
caller-side recoding rules therefore remain selective and do not flatten
refusals, patch conflicts, source changes, or limits.

The ODP doctype safety test now accepts the distinct eager and source-backed
ODF error types through one generic assertion helper. The XLSX raw aggregate
test uses the public slicer codec, checks the package writer through its
imported name, and obtains aggregate input bytes through the metadata
`payload_len()` projection. That projection remains cold and does not decode
deferred payloads before the raw admission check. Existing byte assertions
continue to use `Part::blob()` only after a payload-bearing `get_part()` call.

The v8 all-target diagnostic also exposed test-only imports missing from the
ODS facade, document policy, and DDE tests. The repairs add those local error
aliases and keep the existing assertions for `SourceChanged`, invalid format,
resource limits, typed limits, and patch conflicts.

## Validation

- `rustfmt --edition 2024` passed on all eight changed Rust files.
- `git diff --check` passed.
- The v8 compiler log is retained outside the checkout at
  `(deleted-coord-artifact)/check-0.log`
  (34,232 bytes, SHA-256
  `8144ccd0e597790bc18a871362ac5bb69178d1323f13c3f242457b16c3a5422a`).
- Its reduced diagnostic manifest is
  `compiler-errors.json` (6,518 bytes, SHA-256
  `5396dcf04a83b4ff6f67433e1f15a84fb38595d4be318baba84a4984c5224037`).
- Cargo was not run in this isolated source-only lane. The coordinator will
  rerun the workspace check and affected semantic tests after integration.
