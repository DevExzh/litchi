# ADR 0014: Core-properties reader ownership in the OOXML common crate

- Status: Accepted
- Date: 2026-08-03

## Context

The `litchi-ooxml` migration host contained a complete OPC core-properties
reader even though the grammar is shared by WordprocessingML,
PresentationML, SpreadsheetML, and XLSB packages. Four umbrella-facade call
sites therefore reached through the host for a service that has no vertical
format semantics. The host module also flattened its failures and retained two
production `expect` calls after validation.

Core-properties authoring already belongs to `litchi-ooxml-common`. That crate
also owns the OPC, XML, datetime, decoding, and metadata dependencies required
by the reader, so moving the read grammar adds no dependency edge.

## Decision

`litchi_ooxml_common::properties::read(&OpcPackage)` is the single public read
entry beside `DocumentProperties`. Its implementation remains private under
`properties::read`; the crate root does not add an ambiguous `read` export.

The reader selects the part only through the package-level core-properties
relationship, accepts Transitional and Strict relationship/root namespaces,
enforces OPC M4 cardinality and markup restrictions, validates the declared
content type, bounds each retained property value, and returns structured
common errors for OPC, relationship, content-type, XML, decoding, missing-part,
invalid-value, and resource-limit failures. No validated invariant is recovered
with `expect` or a panic.

The migration-host module and API are deleted without a compatibility alias.
Document, presentation, XLSX, and XLSB facade adapters call the common owner
directly. This ownership move preserves their current policy for optional
metadata failures; deciding whether every malformed optional property is fatal
belongs to the facade validation/report migration rather than this crate cut.

## Consequences

- One host-neutral crate owns both core-properties read and authoring grammar.
- The umbrella no longer routes four metadata reads through
  `litchi-ooxml`; the remaining umbrella-to-host debt is reduced in source
  scope without falsely declaring the whole temporary dependency resolved.
- Removing the host module is intentionally breaking and leaves no forwarding
  shim.
- Package bytes and writers are unchanged, so this decision makes no native
  Office or performance claim.

## Verification

Focused tests cover relationship-selected lookup, absence, external and
dangling targets, wrong content types, duplicate relationships, all modeled
fields, time-zone normalization, writer/parser round trips, DTD and entity
rejection, text budgets, and Apache POI OPC conformance fixtures.
All 13 focused properties tests pass. Warning-denied Clippy and rustdoc are
green for the common owner, migration host, and isolated OOXML facade, as are
workspace lint and the executable boundary policy (35 packages, 107 direct
internal dependencies, and 14 explicit debt items). The previously green
full-workspace test suite is not repeated.
