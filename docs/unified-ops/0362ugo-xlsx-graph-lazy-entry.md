# 0362ugo — XLSX graph-family lazy edit entry

This slice closes the XLSX construction-cost rows for metadata, calculation
properties, connections, Custom Data, external links, the Data Model and
Custom XML Maps.  It starts from `8bf4833ae51153791d61b6be4be54310cf71b142`
on `uo2/xlsx-graphs-lazy`.  Coordinator integration and the final workspace
gate remain pending; this record makes no native timing or performance claim.

## Entry-state design

Each `Package::edit_*` constructor now performs only the existing mutation
policy check and retains an O(1) `record::Source` clone together with the
selected family limits or conformance.  The graph snapshot, source proofs,
relationships and semantic draft are loaded by a retry-safe `ensure_loaded`
path on the first getter, mutation or commit.  A failed load leaves the edit
unloaded, so managed-source, XML, relationship and typed-limit failures are
returned without retaining a partial graph or draft and a later operation can
retry.

Metadata defers the common-properties snapshot and owner admission until the
first metadata operation.  Calculation metadata keeps the workbook limits
until its first semantic access.  Connections, Custom Data, external links,
Data Model and Custom XML Maps similarly defer graph parsing, source capture,
semantic draft cloning and next-identity discovery.  Existing source limits
are still checked before the corresponding snapshot retains payloads, and the
publication paths retain their stale-source, signed-package, exact no-op,
patch, inverse and typed-error behavior.

The fallible first-access signatures are migrated through the XLSX crate,
XLSX ops metadata route, perf-baseline calculation route, raw transaction
tests and format-family tests.  No consumer uses a panic or a placeholder
snapshot to hide deferred errors.

## Validation boundary

No Cargo command was run in this lane because the coordinator had no build
slot during implementation.  Before integration, run the pinned offline,
locked XLSX all-target/all-feature check and the focused graph-family tests
under `(deleted-coord-artifact)`; record exact source and external log
receipts in the follow-up review.  `rustfmt` and `git diff --check` were run
on the edited sources.  No binaries, large evidence files or `3rdparty/`
references were added.
