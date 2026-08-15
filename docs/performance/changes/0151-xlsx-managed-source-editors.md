# Change 0151: managed XLSX source-backed editors

Date: 2026-08-16

Status: production correctness and resource-accounting freeze; no performance
result is claimed.

This change freezes the broad managed source-backed XLSX editor tranche. It
adds managed-package constructors and preserves the same source-backed edit
contracts for these eleven focused editors:

- calculation properties;
- defined names;
- tab state;
- print options;
- page breaks;
- page margins;
- page setup;
- sheet protection;
- data validation;
- auto filter; and
- conditional formatting.

## Ownership and publication boundary

The private `SourcePayload` owner is either `Managed(PartData)` or
`Owned(Arc<Vec<u8>>)`:

- a managed `PartData` handle remains attached to the caller's OPC cache and
  its `Budget` reservation for the lifetime of the snapshot/commit;
- an ordinary materialized package uses the compatibility `Owned(Arc)` path;
  no managed reservation is implied; and
- asking a managed payload to escape into an owning `Arc` is fallible and
  returns the typed `OpcError::ManagedPartDataArcEscape` error. It never
  silently detaches a reservation.

Each editor now exposes the managed `ReadAt`/`SourceCacheLimits`/
`ExecutionContext` constructor combinations and a validated
`SourceBackedPackage` handoff. The handoff checks execution state around
snapshot construction and publication. Source-backed commits consume the
bounded direct OPC publisher: a proven one-Part edit materializes only its
selected Part and raw-copies the rest. The tab-state closure may additionally
publish the workbook and the changed active/visibility worksheet Parts; it
retains the same selected-Part and raw-copy boundary.

## Correctness and resource gates

The focused production tests retain these protections across the eleven
closures:

- an exact semantic no-op publishes the original bytes, including a signed
  source; a changed signed source is refused before output;
- selected Markup Compatibility (MCE), DTD/processing-instruction, protected,
  relationship-changing, unsupported, and unknown-owner/member layouts are
  refused before output. Unselected Parts, relationships, media, and other
  opaque members remain outside the selected closure and are raw-copied when
  the source layout permits the edit;
- source-version changes, foreign commits, stale owner relationships, and
  mismatched content/Part identity are rejected atomically;
- cancellation is checked before stream publication and sink failures remain
  fail-closed; and
- changed output is reopened and checked for typed semantics, source
  preservation, relationship identity, inverse behavior where supported, and
  exact output/no-output behavior.

Managed representative tests charge the exact retained source-Part bytes to
`Resource::Memory`, verify `budget_managed` diagnostics, release the charge
after the snapshot/commit is dropped, and assert that an exact one-byte-under
budget fails before the selected payload is retained and leaves the budget at
zero. Cancellation and selected publication tests likewise require zero
output on refusal and zero retained memory after release. The managed
`PartData` path is therefore resource-accounted correctness evidence, not a
claim that all parsed stores, staging, rewritten candidates, or sink buffers
fit inside that budget.

The validation run for this freeze reported 765 XLSX unit, integration, and
documentation tests green, including 74 focused source-editor checks. No
benchmark selector or performance artifact was added by this change.

## Evidence boundary

This record claims correctness, preservation, typed refusal, and managed
payload retention/release only. It does not claim latency, allocation counts,
RSS or peak memory, copied bytes, decompressed/recompressed bytes, cold I/O,
total-memory bounds, hardware scaling, or real-producer breadth. The tranche
does not add iWork coverage.
