# Change 0304: XLSB source-backed text traversal

Status: implemented

Source-backed XLSB text walks the complete tab catalog in workbook order,
materializes one supported semantic worksheet locally, emits its dense
rectangle, and drops it before processing the next tab. Source and execution
state are checked around materialization, row traversal, sink writes, and
completion. Unsupported tabs return a typed `UnsupportedFeature` error from
the owner instead of silently falling back.

The facade keeps the owner result private while attempting source-backed text.
Every `UnsupportedFeature` variant from the source owner triggers the facade's
only fallback path; every other source error propagates. The eager fallback
parses each compatibility worksheet once and publishes it through the existing
per-position `OnceLock` under its initialization mutex, so repeated and
concurrent operations reuse the published worksheet. Successful source-backed
text does not populate that eager worksheet cache.

Each row is counted without allocation before transient row construction. A
small row that exceeds the configured output limit is still passed to the
shared writer so it returns its normal `Limit` error and progress. If a row
exceeds the bounded limit-probe size, the owner returns a typed document
capacity error before allocation because the shared writer exposes no public
constructor for a synthetic `Limit` value.

performance_claim: none

The only residency claim is that one supported semantic worksheet is resident
at a time on a successful source-backed text traversal. This excludes retained
OPC payload caches, shared-string and style caches, the output `String`, and
unsupported fallback cancellation or eager memory. No RSS, latency, I/O,
allocation, or OOM-proof claim is made. Mixed-tab worksheet indexing is not
claimed to be fixed by this change.
