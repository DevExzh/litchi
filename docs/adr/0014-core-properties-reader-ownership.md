# ADR 0014: Core-properties reader ownership in the OOXML common crate

- Status: Amended by [ADR 0015](0015-lossless-core-properties-crud.md)
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

At this decision's original boundary,
`litchi_ooxml_common::properties::read(&OpcPackage)` was the single public read
entry beside `DocumentProperties`. Its implementation remained private under
`properties::read`; the crate root did not add an ambiguous `read` export. ADR
0015 replaces that value with `Props` and adds the common write and clear
operations without changing reader ownership.

The reader selects the part only through the package-level core-properties
relationship, accepts Transitional and Strict relationship/root namespaces,
enforces OPC M4 cardinality and markup restrictions, validates the declared
content type, bounds each retained property value, and returns structured
common errors for OPC, relationship, content-type, XML, decoding, missing-part,
invalid-value, and resource-limit failures. No validated invariant is recovered
with `expect` or a panic.

The migration-host module and API are deleted without a compatibility alias.
At this decision's original boundary, document, presentation, XLSX, and XLSB
facade adapters called the common owner directly. ADR 0015 subsequently made
malformed core properties fatal at host construction, retained the validated
value in a mutation-tracked host slot for DOCX, PPTX, and XLSX, and reused that
cache at the umbrella seam. XLSB continues to call the common owner directly.

## Consequences

- One host-neutral crate owns both core-properties read and authoring grammar.
- At this decision's original boundary, the umbrella no longer routed four
  metadata reads through `litchi-ooxml`; the remaining umbrella-to-host debt
  was reduced in source scope without falsely declaring the whole temporary
  dependency resolved. ADR 0015 later routes DOCX, PPTX, and XLSX through their
  validated host caches while leaving XLSB on the common reader.
- Removing the host module is intentionally breaking and leaves no forwarding
  shim.
- This ownership-only change made no native Office or performance claim. ADR
  0015 owns the later writer and native-verification evidence.

## Verification

At ADR 0014 acceptance, focused tests covered relationship-selected lookup,
absence, external and dangling targets, wrong content types, duplicate
relationships, all then-modeled fields, time-zone normalization,
writer/parser round trips, DTD and entity rejection, text budgets, and Apache
POI OPC conformance fixtures. All 13 focused properties tests passed.
Warning-denied Clippy and rustdoc were green for the common owner, migration
host, and isolated OOXML facade, as were workspace lint and the executable
boundary policy (35 packages, 107 direct internal dependencies, and 14
explicit debt items). The previously green full-workspace test suite was not
repeated.
