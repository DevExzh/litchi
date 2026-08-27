# Change 0335: OPC source-backed topology-aware Part removal

## Scope

This change adds `SourceTopologyPlan::try_remove_part`, a source-backed OPC
removal plan that composes physical Part omission with the graph edits needed
to keep the remaining package topology valid. It is the graph-safe layer above
the physical ZIP-member deletion primitive from Change 0334: the removed Part
payload is not materialized, while the source relationship and content-type
metadata needed to publish a valid package is inspected and edited.

The plan resolves the target through the immutable source catalog and plans the
whole removal before writing to the sink. The caller must explicitly remove or
retarget every retained inbound internal relationship; publication refuses a
plan that leaves any such edge targeting the removed Part.
It also omits the relationship member owned by the removed Part, and removes
the matching `[Content_Types].xml` `Override` lexically. Unrelated relationship
edges, content-type entries, ZIP records, and source bytes remain eligible for
raw preservation. Format-specific indexes and references inside other Part
payloads are outside this OPC graph operation.

## Implementation

- `try_remove_part` rejects missing, duplicate, case-equivalent, and otherwise
  ambiguous target selections before any sink output.
- Inbound relationships are resolved from package-level and Part-level
  `.rels` owners. A retained owner requires an explicit relationship remove or
  retarget operation in the same plan; its unrelated relationships remain.
- The removed Part's owned `.rels` member is omitted as part of the same
  publication. The operation does not leave a relationship member owned by a
  Part that no longer exists.
- Matching content-type `Override` entries for the Part and for an omitted
  owned `.rels` member are removed by lexical PartName. Defaults, unrelated
  Overrides, ordering, and source spelling outside those entries are not
  regenerated wholesale.
- The removed payload is omitted without reading or decompressing it. Where
  the source layout permits, untouched compressed payloads, local records,
  extras, comments, ordering, central records, and archive framing are copied
  from the source; only the required relationship and content-type records are
  rewritten.
- All graph decisions and selected source members are validated before output,
  so a refusal cannot publish a semantically partial removal. Publication
  remains consuming and sequential, with the existing partial-sink behavior.

## Source, signature, and limit boundaries

This is a source-backed operation over an immutable source catalog. The source
version is checked at the established publication boundaries, and source
changes, cancellation, execution-context failures, sink I/O failures, and
partial output retain the existing typed error semantics. A changed signed
source is refused before output rather than silently publishing a package whose
signature no longer describes its contents. An unchanged/no-op source remains
eligible for the byte-identical source path.

The package's configured `ReadLimits` bound content-type and relationship XML,
member metadata, graph traversal, and materialized planning state. Output
budgets and execution limits are applied to the sequential publisher. Missing
targets, ambiguous catalogs, malformed or unsafe relationship/content-type
metadata, unsupported package topology, and unsupported ZIP layouts remain
typed refusals; there is no eager fallback that would drop preservation or
safety guarantees. ZIP64, trailing-byte, multi-disk, prefixed, overlapping,
and otherwise ambiguous layouts retain the refusal boundaries established by
the source-backed publisher.

## Caller contract

This plan owns OPC topology cleanup only. A format owner must still update
format-specific indexes and references, and must decide whether removing a
Part is semantically valid for that format. The operation does not inspect
arbitrary XML or binary payload references, infer formula dependencies, or
approximate unsupported edits. A caller that needs to remove a target and
repair non-OPC references must compose those changes transactionally with this
publication.

## Accounting and residuals

Operation accounting remains conservative. Accepted output includes ZIP
framing and generated relationship/content-type records; raw unchanged-source
bytes are tracked as archive bytes rather than decompressed payload bytes. The
preflight output bound does not credit omitted local spans, so a source close to
the ZIP32 or output limit can be refused even when the final archive would be
smaller after removal. The counters therefore describe observed and accepted
work, not a benchmark or an exact estimate of the bytes saved by omission.

The changed `.rels` and `[Content_Types].xml` members cannot be treated as
fully raw-preserved records, although bytes unrelated to the selected lexical
edits remain preserved. Format-specific payload cleanup, shared-Part ownership
policy, and any semantic references not represented by OPC relationships remain
residuals for the owning crate.

## Validation evidence

All Cargo commands used Rust 1.95.0, `CARGO_BUILD_JOBS=1`, and the isolated
`/dev/shm/litchi-0335-target` directory:

- Focused `source_backed_topology` suite: 28 passed.
- Complete `litchi-opc` library suite: 264 passed.
- Complete `litchi-opc` integration suites: 69 passed.
- `cargo +1.95.0 clippy -p litchi-opc --lib --tests -- -D warnings`: passed.
- `RUSTDOCFLAGS="-D warnings" cargo +1.95.0 doc -p litchi-opc --no-deps`:
  passed.
- `cargo +1.95.0 fmt -p litchi-opc -- --check`: passed before commit.

## Performance claims

`performance_claim: none`

No benchmark, latency, throughput, RSS, allocation, or process-wide memory
claim is made. Avoiding the removed payload read and preserving untouched ZIP
spans are architectural properties of this path, not measured performance
results.

## Follow-up

This closes the `calcChain` publication seam for future source-backed XLSX
formula transactions. The XLSX owner can use one topology publication to omit
`xl/calcChain.xml`, detach its workbook relationship, omit its owned
relationship member when applicable, and remove its content-type Override,
while retaining the source, signature, limit, and preservation guarantees
above. XLSX remains responsible for calculation-chain and formula semantics
outside the OPC graph.
